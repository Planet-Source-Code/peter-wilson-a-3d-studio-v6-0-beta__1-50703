Attribute VB_Name = "m3DMaths"
Option Explicit

' =========================================================================================
' 3D Computer Graphics for Visual Basic Programmers: Theory, Practice, Source Code and Fun!
' Version: 6.0 beta - Precision Edition
'
' by Peter Wilson
' Copyright © 2004 - Peter Wilson - All rights reserved.
' http://dev.midar.com/
' =========================================================================================


' Define the name of this class/module for error-trap reporting.
Private Const m_strModuleName As String = "m3DMaths"

' =========================================================================================
' Define a few constants.
' =========================================================================================
Private Const g_sngPI As Double = 3.14159265358979
Private Const g_sngPIDivideBy180 As Double = 1.74532925199433E-02
Private Const g_sng180DivideByPI As Double = 57.2957795130823


Public Function MatrixShear(ShearX As Single, ShearY As Single) As mdrMatrix4
    
    ' Create a new Identity matrix (i.e. Reset)
    MatrixShear = MatrixIdentity()
    
    ' Shear along the X and Y axis.
    '
    ' Shearing is used to distort an image for a very particular purpose. More specifically, it's
    ' used to help correctly reorient both the 3D objects and the observer (you)
    ' so that the image does NOT look distorted.
    ' You will NOT need to apply the Shear matrix to any part of your 3D objects (planes, tanks, etc.)
    
    MatrixShear.rc13 = ShearX
    MatrixShear.rc23 = ShearY
    
End Function

Public Function MatrixViewMapping_Per(p_Camera As mdr3DTargetCamera) As mdrMatrix4
        
    Dim vectCW As mdrVector4            '   Centre of Window
    Dim vectDOP As mdrVector4           '   Direction Of Projection
    Dim matTranslate As mdrMatrix4
    
    Dim sngShearX As Single
    Dim sngShearY As Single
    Dim matShear As mdrMatrix4
    
    Dim sngScaleX As Single
    Dim sngScaleY As Single
    Dim sngScaleZ As Single
    Dim matScale As mdrMatrix4
    
    Dim matPerspective As mdrMatrix4
    
    
    ' Translate such that the centre of projection (COP), given by PRP, is at the origin (p. 268)
    ' ===========================================================================================
    matTranslate = MatrixTranslation(-p_Camera.PRP.x, -p_Camera.PRP.y, -p_Camera.PRP.Z)
    
    
    ' Calculate the Centre of the Window.
    ' ===================================
    vectCW.x = (p_Camera.Umax + p_Camera.Umin) / 2
    vectCW.y = (p_Camera.Vmax + p_Camera.Vmin) / 2
    vectCW.Z = 0
    vectCW.w = 1
    
    
    ' Calculate the difference between the Centre of the Window, and the PRP.
    ' The result is the Direction Of Projectsion (DOP), which should be the opposite of VPN.
    ' The DOP points in the direction of the camera, ie. Direction of Projection.
    ' ======================================================================================
    vectDOP = VectorSubtract(vectCW, p_Camera.PRP)
    
    
    ' Calculate the Shear Matrix
    ' ==========================
    If vectDOP.Z <> 0 Then
        sngShearX = -(vectDOP.x / vectDOP.Z)
        sngShearY = -(vectDOP.y / vectDOP.Z)
    End If
    matShear = MatrixShear(sngShearX, sngShearY)
    
    
    ' Calculate the Perspective Scale transformation, such that the view volume becomes the
    ' canonical perspective view volume, the truncated right pyramid defined by the
    ' six planes (ready for clipping) Eq. 6.39 on p.269.
    ' =======================================================================================
    Dim sngTemp As Double
    sngScaleX = (2 * -p_Camera.PRP.Z) / ((p_Camera.Umax - p_Camera.Umin) * (-p_Camera.PRP.Z + p_Camera.ClipFar))
    sngScaleY = (2 * -p_Camera.PRP.Z) / ((p_Camera.Vmax - p_Camera.Vmin) * (-p_Camera.PRP.Z + p_Camera.ClipFar))
    sngScaleZ = -1 / (-p_Camera.PRP.Z + p_Camera.ClipFar)
    matScale = MatrixScale(sngScaleX, sngScaleY, sngScaleZ)
    
    
    ' Ok... now that we have the "perspective-projection canonical view volume" (above), it is normal to
    ' covert this into the "parallel-projection canonical view volume". This is so a single clipping procedure
    ' can be used for both perspective and parallel.
    '
    ' zMin is the transform front clipping plane (Eq. 6.48)
    Dim sngZmin As Double
    sngZmin = -((-p_Camera.PRP.Z + p_Camera.ClipNear) / (-p_Camera.PRP.Z + p_Camera.ClipFar))
    matPerspective = MatrixIdentity
    If sngZmin <> -1 Then ' Minus one is the only value not allowed!
        matPerspective.rc33 = 1 / (1 + sngZmin)
        matPerspective.rc34 = -sngZmin / (1 + sngZmin)
        matPerspective.rc43 = -1
        matPerspective.rc44 = 0
    End If
    
    
    MatrixViewMapping_Per = MatrixIdentity()
    MatrixViewMapping_Per = MatrixMultiply(MatrixViewMapping_Per, matTranslate)
    MatrixViewMapping_Per = MatrixMultiply(MatrixViewMapping_Per, matShear)
    MatrixViewMapping_Per = MatrixMultiply(MatrixViewMapping_Per, matScale)
    MatrixViewMapping_Per = MatrixMultiply(MatrixViewMapping_Per, matPerspective)
    
End Function

Public Function Matrix_vv3dv(Xmin As Single, Xmax As Single, Ymin As Single, Ymax As Single, Zmin As Single, Zmax As Single, Optional blnKeepSquare As Boolean = True) As mdrMatrix4
    
    Dim matTranslateA As mdrMatrix4
    Dim matTranslateB As mdrMatrix4
    Dim sngAspectRatio As Single
    Dim sngScaleX As Single
    Dim sngScaleY As Single
    Dim sngScaleZ As Single
    Dim matScale As mdrMatrix4
    
    Dim sngDelta1 As Single
    
    
    ' Translate the canonical parallel projection view volume so that it's corner (-1,-1,-1) becomes
    ' the origin. This is so that the scaling process does not distort the geometry.
    ' ==============================================================================================
    matTranslateA = MatrixTranslation(1, 1, 1)
    
    
    ' The translated view volume is scaled into the size of the 3D viewport, with the following scale.
    ' This is the part where you turn those 'virtual' camera coordinates into real screen pixels!
    ' This scale matrix can also flip the y-coordinates so that (0,0) is at the bottom-left instead
    ' of the default windows location at top-left. You can also adjust the aspect ratio of the output here.
    ' NOTE for VB users: Visual Basic lets you change the scale-mode setting of a form,picturebox or pinter,
    '   This routine does pretty much the same thing, so if you wanted you could skip this
    '   "Matrix_vv3dv" step altogether and simply adjust the form/picturebox scale settings instead.
    '   Either you "scale the geometry" to "fit the window", or you "scale the window" to "fit the geometry".
    '   You would use something like: myForm.ScaleLeft=myCamera.Umin: myForm.ScaleWidth=myCamera.Umax, etc.
    '   PS.
    '   I think this is correct, you may need to double-check. The term, "scaling the geometry" sounds wrong.
    ' =======================================================================================================
    
    sngScaleX = (Xmax - Xmin) / 2
    sngScaleY = (Ymax - Ymin) / 2
    sngScaleZ = (Zmax - Zmin) / 1
    
    If blnKeepSquare = True Then
    
        ' Keeps the image square and centered within the window. It also flips the image so that (0,0)
        ' is at the bottom-left (which is different to normal windows behaviour with (0,0) at top-left).
        ' ==============================================================================================
        sngAspectRatio = Abs((Xmax - Xmin) / (Ymax - Ymin))  ' X pixels are sngAspectRatio times bigger than Y pixels.
        matScale = MatrixScale(sngScaleX, sngScaleY * sngAspectRatio, sngScaleZ)
        sngDelta1 = (Xmax - Ymin) / 2
        
    Else
    
        ' Assume the window's pixel origin of (0,0) is at the bottom-left of the screen.
        ' I am unaware of OS that does this, but I leave this in to be consistent with graphics literature.
        ' ==================================================================================================
        matScale = MatrixScale(sngScaleX, sngScaleY, sngScaleZ)
        sngDelta1 = 0 ' just ignore
        
    End If
    
    
    ' Finally, the properly scaled view volume is translated
    ' to the lower-left corner of the viewport
    ' =======================================================
    matTranslateB = MatrixTranslation(Xmin, Ymin + sngDelta1, Zmin)
    
    
    ' Section: 6.5.5
    Matrix_vv3dv = MatrixIdentity()
    Matrix_vv3dv = MatrixMultiply(Matrix_vv3dv, matTranslateA)
    Matrix_vv3dv = MatrixMultiply(Matrix_vv3dv, matScale)
    Matrix_vv3dv = MatrixMultiply(Matrix_vv3dv, matTranslateB)

End Function
Public Function MatrixScale(ScaleX As Single, ScaleY As Single, ScaleZ As Single) As mdrMatrix4
    
    ' Create a new Identity matrix (i.e. Reset)
    MatrixScale = MatrixIdentity()
    
    ' Makes an object bigger or smaller on any of the three axes.
    ' Note: Some imported .x files need scaling in the order of 100, 200 even 1000, while other objects
    '       do not need scaling at all.
    '       ie. If the scale factor is 2,2,2 then the object is doubled in size on all three
    '           axes. A scale factor of 0.5,0.5,0.5 will shrink an object on all three axes.
    '
    ' Normally you scale on all three axis, by the same amount, otherwise your 3D object may get
    ' distorted in a way you didn't expect. If in doubt, just make all the numbers the same....
    ' a Uniform Scale.
    
    MatrixScale.rc11 = ScaleX
    MatrixScale.rc22 = ScaleY
    MatrixScale.rc33 = ScaleZ
    
End Function

Public Function ConvertDeg2Rad(Degress As Single) As Single
Attribute ConvertDeg2Rad.VB_Description = "Converts Degrees to Radians."

    ' Converts Degrees to Radians
    ConvertDeg2Rad = Degress * (g_sngPIDivideBy180)
    
End Function

Public Function ConvertRad2Deg(Radians As Single) As Single
Attribute ConvertRad2Deg.VB_Description = "Converts Radians to Degrees."
 
    ' Converts Radians to Degrees
    ConvertRad2Deg = Radians * (g_sng180DivideByPI)
    
End Function

Public Function MatrixViewOrientation(vectVPN As mdrVector4, vectVUP As mdrVector4, vectVRP As mdrVector4) As mdrMatrix4
Attribute MatrixViewOrientation.VB_Description = "Builds a ViewOrientation Matrix to correctly orientate the scene. VPN = View Plane Normal, VUP=Up Vector, VRP=View Reference Point."
    
    ' =====================================================
    ' Rotate VRC such that the:
    '   * n axis becomes the z axis,
    '   * u axis becomes the x axis and
    '   * v axis becomes the y axis.
    ' =====================================================
    
    Dim matRotateVRC As mdrMatrix4
    Dim matTranslateVRP As mdrMatrix4
    
    Dim vectN As mdrVector4
    Dim vectU As mdrVector4
    Dim vectV As mdrVector4
    
        
    '         VPN
    ' n* = ¯¯¯¯¯¯¯¯¯¯¯
    '       | VPN |
    '
    ' * also referred to as Rz (eq. 6.25)
    vectN = VectorNormalize(vectVPN)
    
    
    '       VUP x n
    ' u* = ¯¯¯¯¯¯¯¯¯¯¯¯¯
    '     | VUP x n |
    '
    ' * Also referred to as Rx (eq.6.26)
    vectU = CrossProduct(vectVUP, vectN)
    vectU = VectorNormalize(vectU)
    
    
    ' v* = n x u
    '
    ' * Also referred to as Ry (eq.6.27)
    vectV = CrossProduct(vectN, vectU)
    
    
    ' Define the Rotate matrix such that the n-axis (VPN) becomes the z-axis,
    ' the u-axis becomes the x-axis and the v-axis becomes the y-axis.
    matRotateVRC = MatrixIdentity()
    With matRotateVRC
        .rc11 = vectU.x: .rc12 = vectU.y: .rc13 = vectU.Z
        .rc21 = vectV.x: .rc22 = vectV.y: .rc23 = vectV.Z
        .rc31 = vectN.x: .rc32 = vectN.y: .rc33 = vectN.Z
    End With
    
    
    ' Define a Translation matrix to transform the VRP to the origin.
    matTranslateVRP = MatrixTranslation(-vectVRP.x, -vectVRP.y, -vectVRP.Z)
    
    
    ' Theory
    ' ===============================================================================
    ' MatrixViewOrientation =  matTranslateVRP * matRotateVRC
    '                          (Remember, read this and calculate from Right to Left)
    ' ===============================================================================
    MatrixViewOrientation = MatrixIdentity()
    MatrixViewOrientation = MatrixMultiply(MatrixViewOrientation, matTranslateVRP)
    MatrixViewOrientation = MatrixMultiply(MatrixViewOrientation, matRotateVRC)
    
    
End Function

Public Function VectorSubtract(V1 As mdrVector4, v2 As mdrVector4) As mdrVector4
Attribute VectorSubtract.VB_Description = "Returns the result of Vector2 subtracted from Vector1."

    ' Subtracts vector 2 away from vector 1.
    With VectorSubtract
        .x = V1.x - v2.x
        .y = V1.y - v2.y
        .Z = V1.Z - v2.Z
        .w = 1 ' Ignore W
    End With
    
End Function

Public Function MatrixTranspose(MIn As mdrMatrix4) As mdrMatrix4
    
    ' Swaps Rows for Columns (and visa-versa) in a 4x4 matrix.
    
    With MatrixTranspose
        
        .rc11 = MIn.rc11: .rc12 = MIn.rc21: .rc13 = MIn.rc31: .rc14 = MIn.rc41
        .rc21 = MIn.rc12: .rc22 = MIn.rc22: .rc23 = MIn.rc32: .rc24 = MIn.rc42
        .rc31 = MIn.rc13: .rc32 = MIn.rc23: .rc33 = MIn.rc33: .rc34 = MIn.rc43
        .rc41 = MIn.rc14: .rc42 = MIn.rc24: .rc43 = MIn.rc34: .rc44 = MIn.rc44
        
    End With
    
End Function

Public Function VectorAddition(V1 As mdrVector4, v2 As mdrVector4) As mdrVector4
Attribute VectorAddition.VB_Description = "Returns the result of two Vectors added together."

    ' Adds two vectors together.
    With VectorAddition
        .x = V1.x + v2.x
        .y = V1.y + v2.y
        .Z = V1.Z + v2.Z
        .w = 1 ' Ignore W
    End With
    
End Function

Public Function MatrixTranslation(OffsetX As Single, OffsetY As Single, OffsetZ As Single) As mdrMatrix4
Attribute MatrixTranslation.VB_Description = "Given X, Y and Z offsets, builds a Translation Matrix."
    
    ' Translation is another word for "move".
    ' ie. You can translate an object from one location to another.
    '     You can    move   an object from one location to another.
    '
    ' The ability to combine a Rotation with a Translation within a single matrix, is the main
    ' reason why I have used a 4x4 matrix and NOT a 3x3 matrix.
    
    ' Create a new Identity matrix (i.e. Reset)
    MatrixTranslation = MatrixIdentity()
    
    With MatrixTranslation
        .rc14 = OffsetX
        .rc24 = OffsetY
        .rc34 = OffsetZ
    End With
    
    ' Very important note about this matrix
    ' =====================================
    ' If you see other programmers placing their Offset's in different positions (like the columns
    ' and rows have been swapped over - ie. Transposed) then this probably means that they have coded all
    ' of their algorithims to a different "notation standard". This subroutine follows the conventions used
    ' in the ledgendary bible "Computer Graphics Principles and Practice", Foley·vanDam·Feiner·Hughes which
    ' illustrates mathematical formulas using Column-Vector notation. Other books like "3D Math Primer for
    ' Graphics and Game Development", Fletcher Dunn·Ian Parberry, use Row-Vector notation. Both are correct,
    ' however it's important to know which standard you code to, because it affects the way in which you
    ' build your matrices and the order in which you should multiply them to obtain the correct result.
    '
    ' OpenGL uses Column Vectors (like this application).
    ' DirectX uses Row Vectors.
    
End Function

Public Function MatrixIdentity() As mdrMatrix4
Attribute MatrixIdentity.VB_Description = "Returns an Identity matrix."

    ' The identity matrix is used as the starting point for matrices
    ' that will modify vertex values to create rotations, translations,
    ' and any other transformations that can be represented by a 4×4 matrix
    '
    ' Notice that...
    '   * the 1's go diagonally down?
    '   * rc stands for Row Column. Therefore, rc12 means Row1, Column 2.
    '
    ' Comments:
    ' You'll often hear people talking about the "identity matrix"... well this is it!
    ' Sometimes the identify matrix also contains pre-calculated rotations and translations. This is usually
    ' the case when you import a 3D object from another application.
    
    With MatrixIdentity
        .rc11 = 1: .rc12 = 0: .rc13 = 0: .rc14 = 0
        .rc21 = 0: .rc22 = 1: .rc23 = 0: .rc24 = 0
        .rc31 = 0: .rc32 = 0: .rc33 = 1: .rc34 = 0
        .rc41 = 0: .rc42 = 0: .rc43 = 0: .rc44 = 1
    End With
    
End Function

Public Function MatrixMultiply(m1 As mdrMatrix4, m2 As mdrMatrix4) As mdrMatrix4
Attribute MatrixMultiply.VB_Description = "Returns the result of Matrix1 multiplied by Matrix2."
    
    ' Re-declare m1 & m2
    Dim m1b As mdrMatrix4
    Dim m2b As mdrMatrix4
    m1b = m1
    m2b = m2
    
    ' Matrix multiplication is a set of "dot products" between the rows of the left matrix and columns of the right matrix.
    '
    ' Matrix A and B below
    ' ====================
    '                          | a, b, c |       | j, k, l |
    '  Let A*B represent...    | d, e, f |   *   | m, n, o |
    '                          | g, h, i |       | p, q, r |
    '
    '  Multipling out we get...
    '
    '   | (a*j)+(b*m)+(c*p), (a*k)+(b*n)+(c*q), (a*l)+(b*o)+(c*r) |
    '   | (d*j)+(e*m)+(f*p), (d*k)+(e*n)+(f*q), (d*l)+(e*o)+(f*r) |
    '   | (g*j)+(h*m)+(i*p), (g*k)+(h*n)+(i*q), (g*l)+(h*o)+(i*r) |
    '
    ' To put this another way...
    '
    '  | a, b, c |     | j, k, l |     | (a*j)+(b*m)+(c*p), (a*k)+(b*n)+(c*q), (a*l)+(b*o)+(c*r) |
    '  | d, e, f |  *  | m, n, o |  =  | (d*j)+(e*m)+(f*p), (d*k)+(e*n)+(f*q), (d*l)+(e*o)+(f*r) |
    '  | g, h, i |     | p, q, r |     | (g*j)+(h*m)+(i*p), (g*k)+(h*n)+(i*q), (g*l)+(h*o)+(i*r) |
    '
    ' Note: This was only a 3x3 matrix show... however this routine is actually bigger, using a 4x4.
    ' I just wanted to keep the example short.
    
    
    ' =====================
    ' About this subroutine
    ' =====================
    ' This is the kind of routine that is hard coded into the electronic circuts of many CPU's and
    ' all 3D video cards (actually most of this module is hard coded into the video-cards, in some way or another)
    ' For additional research try searching for "Matrix Multiplication"
    '
    ' Multiply two 4x4 matrices (m2 & m1) and return the result in 'MatrixMultiply'.
    '   64 Floating point multiplications
    '   48 Floating point additions
    '
    ' This matrix multiplies a full 4x4 matrix, however some programmers and/or algorithms only
    ' multiply the top-left 3x3; yes, you can do this, however a 4x4 matrix lets you combine rotation
    ' and movement in a single matrix. If you are using a 3x3 matrix then you can't do this and
    ' will have to calculate rotation and movement as separate steps. A 3x3 matrix also makes it
    ' harder to rotate an object around a point that is not it's origin. Heck! There's a lot of
    ' agruments about 3x3 vs. 4x4, and I can't be bothered getting into them. Just do it the correct
    ' way and everyone will be happy! ;-)
    
    
    ' Reset the matrix to identity.
    MatrixMultiply = MatrixIdentity()
    
    
    With MatrixMultiply
        .rc11 = (m1b.rc11 * m2b.rc11) + (m1b.rc21 * m2b.rc12) + (m1b.rc31 * m2b.rc13) + (m1b.rc41 * m2b.rc14)
        .rc12 = (m1b.rc12 * m2b.rc11) + (m1b.rc22 * m2b.rc12) + (m1b.rc32 * m2b.rc13) + (m1b.rc42 * m2b.rc14)
        .rc13 = (m1b.rc13 * m2b.rc11) + (m1b.rc23 * m2b.rc12) + (m1b.rc33 * m2b.rc13) + (m1b.rc43 * m2b.rc14)
        .rc14 = (m1b.rc14 * m2b.rc11) + (m1b.rc24 * m2b.rc12) + (m1b.rc34 * m2b.rc13) + (m1b.rc44 * m2b.rc14)
        
        .rc21 = (m1b.rc11 * m2b.rc21) + (m1b.rc21 * m2b.rc22) + (m1b.rc31 * m2b.rc23) + (m1b.rc41 * m2b.rc24)
        .rc22 = (m1b.rc12 * m2b.rc21) + (m1b.rc22 * m2b.rc22) + (m1b.rc32 * m2b.rc23) + (m1b.rc42 * m2b.rc24)
        .rc23 = (m1b.rc13 * m2b.rc21) + (m1b.rc23 * m2b.rc22) + (m1b.rc33 * m2b.rc23) + (m1b.rc43 * m2b.rc24)
        .rc24 = (m1b.rc14 * m2b.rc21) + (m1b.rc24 * m2b.rc22) + (m1b.rc34 * m2b.rc23) + (m1b.rc44 * m2b.rc24)
        
        .rc31 = (m1b.rc11 * m2b.rc31) + (m1b.rc21 * m2b.rc32) + (m1b.rc31 * m2b.rc33) + (m1b.rc41 * m2b.rc34)
        .rc32 = (m1b.rc12 * m2b.rc31) + (m1b.rc22 * m2b.rc32) + (m1b.rc32 * m2b.rc33) + (m1b.rc42 * m2b.rc34)
        .rc33 = (m1b.rc13 * m2b.rc31) + (m1b.rc23 * m2b.rc32) + (m1b.rc33 * m2b.rc33) + (m1b.rc43 * m2b.rc34)
        .rc34 = (m1b.rc14 * m2b.rc31) + (m1b.rc24 * m2b.rc32) + (m1b.rc34 * m2b.rc33) + (m1b.rc44 * m2b.rc34)
        
        .rc41 = (m1b.rc11 * m2b.rc41) + (m1b.rc21 * m2b.rc42) + (m1b.rc31 * m2b.rc43) + (m1b.rc41 * m2b.rc44)
        .rc42 = (m1b.rc12 * m2b.rc41) + (m1b.rc22 * m2b.rc42) + (m1b.rc32 * m2b.rc43) + (m1b.rc42 * m2b.rc44)
        .rc43 = (m1b.rc13 * m2b.rc41) + (m1b.rc23 * m2b.rc42) + (m1b.rc33 * m2b.rc43) + (m1b.rc43 * m2b.rc44)
        .rc44 = (m1b.rc14 * m2b.rc41) + (m1b.rc24 * m2b.rc42) + (m1b.rc34 * m2b.rc43) + (m1b.rc44 * m2b.rc44)
    End With
    
End Function

Public Function MatrixMultiplyVector(m1 As mdrMatrix4, V1 As mdrVector4) As mdrVector4
Attribute MatrixMultiplyVector.VB_Description = "Returns the result of a Matrix multiplied by a Vector."
        
    ' Here is a Column Vector (having three letters/numbers)...
    '
    '   | a |
    '   | b |
    '   | c |
    '
    ' Here is the Row Vector equivalent...
    '
    '   | a, b, c |
    '
    ' The two different conventions (Column Vector, Row Vector) store exactly the same information,
    ' so the issue of which is best will not even be discussed!  Just remember that different authors use different
    ' conventions, and it's quite easy to get them mixed up with each other!
    
    
    
    ' Matrix multiplication is a set of "dot products" between the rows of the left matrix and columns of the right matrix.
    '
    ' Matrix A and B below
    ' ====================
    '                            | a, b, c |     | x |
    '  Note the following...     | d, e, f |  *  | y |
    '                            | g, h, i |     | z |
    '
    '  ...multipling out we get...
    '
    '   | (a*x)+(b*y)+(c*z) |
    '   | (d*x)+(e*y)+(f*z) |
    '   | (g*x)+(h*y)+(i*z) |
    
    '
    ' Therefore...
    '
    '   | a, b, c |     | x |     | (a*x)+(b*y)+(c*z) |
    '   | d, e, f |  *  | y |  =  | (d*x)+(e*y)+(f*z) |
    '   | g, h, i |     | z |     | (g*x)+(h*y)+(i*z) |
    
    
    
    
    
    ' Multiply two matrices (m1 & v1) and returns the result in VOut.
    '
    ' m1 is a 4x4 matrix (ColumnsN = 4)
    ' v1 is a Column vector matrix (RowsM = 4 rows)
    '
    ' Because ColumnsN equals RowsM, this is considered a 'Square Matrix' and can be multiplied.
    ' (Notice how the reverse is NOT true: Columns of v1 = 1, Rows of m1 = 4, they are not the
    '  same and thus can't be multiplied in reverse order.)
    '
    ' 16 Floating point multiplications
    ' 12 Floating point additions
    
    With MatrixMultiplyVector
        .x = (m1.rc11 * V1.x) + (m1.rc12 * V1.y) + (m1.rc13 * V1.Z) + (m1.rc14 * V1.w)
        .y = (m1.rc21 * V1.x) + (m1.rc22 * V1.y) + (m1.rc23 * V1.Z) + (m1.rc24 * V1.w)
        .Z = (m1.rc31 * V1.x) + (m1.rc32 * V1.y) + (m1.rc33 * V1.Z) + (m1.rc34 * V1.w)
        .w = (m1.rc41 * V1.x) + (m1.rc42 * V1.y) + (m1.rc43 * V1.Z) + (m1.rc44 * V1.w)
    End With
    
End Function

Public Function VectorNormalize(v As mdrVector4) As mdrVector4
Attribute VectorNormalize.VB_Description = "Returns the normalized version of a Vector. The resulting Vector will have a length equal to 1.0"

    ' Returns the normalized version of a vector.
    
    Dim sngLength As Single
    
    sngLength = VectorLength(v)
    If sngLength = 0 Then sngLength = 1
    
    With VectorNormalize
        .x = v.x / sngLength
        .y = v.y / sngLength
        .Z = v.Z / sngLength
        .w = v.w ' Ignore W
    End With
    
End Function

Public Function VectorLength(v As mdrVector4) As Single
Attribute VectorLength.VB_Description = "Returns the length of a Vector using Pythagoras therom."

    ' Returns the length of a Vector.
    '
    ' In Mathematic books, the "length of a vector" is often written with two vertical bars on either
    ' side, like this:  ||v||
    ' It took me ages to figure this out! Nobody explained it, they just assumed I knew it!
    '
    ' The length of a vector is from the origin (0,0,0) to x,y,z
    ' Do you remember high schools maths, Pythagoras theorem?  c^2 = a^2 + b^2
    '   "In a right-angled triangle, the area of the square of the hypotenuse (the longest side)
    '    is equal to the sum of the areas of the squares drawn on the other two sides."
    
    VectorLength = Sqr((v.x ^ 2) + (v.y ^ 2) + (v.Z ^ 2))
    ' Ignore W
    
End Function

Public Function CrossProduct(vectV As mdrVector4, VectW As mdrVector4) As mdrVector4
Attribute CrossProduct.VB_Description = "Returns the CrossProduct of two vectors."

    ' Determines the cross-product of two 3-D vectors (V and W).
    ' The cross-product is used to find a vector that is perpendicular to the plane defined by VectV and VectW.
    
    With CrossProduct
        .x = (vectV.y * VectW.Z) - (vectV.Z * VectW.y)
        .y = (vectV.Z * VectW.x) - (vectV.x * VectW.Z)
        .Z = (vectV.x * VectW.y) - (vectV.y * VectW.x)
        .w = 1 ' Ignore W
    End With
    
End Function

Public Function VectorMultiplyByScalar(VectorIn As mdrVector4, Scalar As Single) As mdrVector4
Attribute VectorMultiplyByScalar.VB_Description = "Returns the result of a Vector multiplied by a scalar value. Useful for making vectors bigger or smaller."
    
    With VectorMultiplyByScalar
        .x = CSng(VectorIn.x) * CSng(Scalar)
        .y = CSng(VectorIn.y) * CSng(Scalar)
        .Z = CSng(VectorIn.Z) * CSng(Scalar)
        .w = VectorIn.w ' Ignore W
    End With
    
End Function

Public Function MatrixRotationX(Radians As Single) As mdrMatrix4
Attribute MatrixRotationX.VB_Description = "Given an angle expressed as Radians, this function builds a Rotation Matrix for the X Axis."

    ' ===========================================================================================
    '               *** This application uses a Right-Handed Coordinate System ***
    ' ===========================================================================================
    '   The positive X axis points towards the right.
    '   The positive Y axis points upwards to the top of the screen.
    '   The positive Z axis points out of the monitor towards you.
    '
    ' Note: DirectX uses a Left-Handed Coordinate system, which many people find more intuitive.
    ' ===========================================================================================
    
    Dim sngCosine As Double
    Dim sngSine As Double
    
    sngCosine = Round(Cos(Radians), 6)
    sngSine = Round(Sin(Radians), 6)
    
    ' Create a new Identity matrix (i.e. Reset)
    MatrixRotationX = MatrixIdentity()
    
    ' =======================================================================================================
    ' Positive rotations in a right-handed coordinate system are such that, when looking from a
    ' positive axis back towards the origin (0,0,0), a 90° "counter-clockwise" rotation will
    ' transform one positive axis into the other:
    '
    ' X-Axis rotation.
    ' A positive rotation of 90° transforms the +Y axis into the +Z axis.
    ' An additional positive rotation of 90° transforms the +Z axis into the -Y axis.
    ' An additional positive rotation of 90° transforms the -Y axis into the -Z axis.
    ' An additional positive rotation of 90° transforms the -Z axis into the +Y axis (back where we started).
    ' =======================================================================================================
    With MatrixRotationX
        .rc22 = sngCosine
        .rc23 = -sngSine
        .rc32 = sngSine
        .rc33 = sngCosine
    End With
    
    ' Very important note about this matrix
    ' =====================================
    ' If you see other programmers placing their Sines and Cosines in different positions (like the columns
    ' and rows have been swapped over - ie. Transposed) then this probably means that they have coded all
    ' of their algorithims to a different "notation standard". This subroutine follows the conventions used
    ' in the ledgendary bible "Computer Graphics Principles and Practice", Foley·vanDam·Feiner·Hughes which
    ' illustrates mathematical formulas using Column-Vector notation. Other books like "3D Math Primer for
    ' Graphics and Game Development", Fletcher Dunn·Ian Parberry, use Row-Vector notation. Both are correct,
    ' however it's important to know which standard you code to, because it affects the way in which you
    ' build your matrices and the order in which you should multiply them to obtain the correct result.
    '
    ' OpenGL uses Column Vectors (like this application).
    ' DirectX uses Row Vectors.
    
End Function

Public Function MatrixRotationY(Radians As Single) As mdrMatrix4
Attribute MatrixRotationY.VB_Description = "Given an angle expressed as Radians, this function builds a Rotation Matrix for the Y Axis."

    ' ===========================================================================================
    '               *** This application uses a Right-Handed Coordinate System ***
    ' ===========================================================================================
    '   The positive X axis points towards the right.
    '   The positive Y axis points upwards to the top of the screen.
    '   The positive Z axis points out of the monitor towards you.
    '
    ' Note: DirectX uses a Left-Handed Coordinate system, which many people find more intuitive.
    ' ===========================================================================================
    
    Dim sngCosine As Double
    Dim sngSine As Double
    
    sngCosine = Round(Cos(Radians), 6)
    sngSine = Round(Sin(Radians), 6)
    
    ' Create a new Identity matrix (i.e. Reset)
    MatrixRotationY = MatrixIdentity()
    
    ' =======================================================================================================
    ' Positive rotations in a right-handed coordinate system are such that, when looking from a
    ' positive axis back towards the origin (0,0,0), a 90° "counter-clockwise" rotation will
    ' transform one positive axis into the other:
    '
    ' Y-Axis rotation.
    ' A positive rotation of 90° transforms the +Z axis into the +X axis
    ' An additional positive rotation of 90° transforms the +X axis into the -Z axis.
    ' An additional positive rotation of 90° transforms the -Z axis into the -X axis.
    ' An additional positive rotation of 90° transforms the -X axis into the +Z axis (back where we started).
    ' =======================================================================================================
    With MatrixRotationY
        .rc11 = sngCosine
        .rc31 = -sngSine
        .rc13 = sngSine
        .rc33 = sngCosine
    End With
    
    ' Very important note about this matrix
    ' =====================================
    ' If you see other programmers placing their Sines and Cosines in different positions (like the columns
    ' and rows have been swapped over - ie. Transposed) then this probably means that they have coded all
    ' of their algorithims to a different "notation standard". This subroutine follows the conventions used
    ' in the ledgendary bible "Computer Graphics Principles and Practice", Foley·vanDam·Feiner·Hughes which
    ' illustrates mathematical formulas using Column-Vector notation. Other books like "3D Math Primer for
    ' Graphics and Game Development", Fletcher Dunn·Ian Parberry, use Row-Vector notation. Both are correct,
    ' however it's important to know which standard you code to, because it affects the way in which you
    ' build your matrices and the order in which you should multiply them to obtain the correct result.
    '
    ' OpenGL uses Column Vectors (like this application).
    ' DirectX uses Row Vectors.

End Function

Public Function MatrixRotationZ(Radians As Single) As mdrMatrix4
Attribute MatrixRotationZ.VB_Description = "Given an angle expressed as Radians, this function builds a Rotation Matrix for the Z Axis."

    ' ===========================================================================================
    '               *** This application uses a Right-Handed Coordinate System ***
    ' ===========================================================================================
    '   The positive X axis points towards the right.
    '   The positive Y axis points upwards to the top of the screen.
    '   The positive Z axis points out of the monitor towards you.
    '
    ' Note: DirectX uses a Left-Handed Coordinate system, which many people find more intuitive.
    ' ===========================================================================================
    
    
    Dim sngCosine As Double
    Dim sngSine As Double
    
    sngCosine = Round(Cos(Radians), 6)
    sngSine = Round(Sin(Radians), 6)
    
    ' Create a new Identity matrix (i.e. Reset)
    MatrixRotationZ = MatrixIdentity()

    ' =======================================================================================================
    ' Positive rotations in a right-handed coordinate system are such that, when looking from a
    ' positive axis back towards the origin (0,0,0), a 90° "counter-clockwise" rotation will
    ' transform one positive axis into the other:
    '
    ' Z-Axis rotation.
    ' A positive rotation of 90° transforms the +X axis into the +Y axis.
    ' An additional positive rotation of 90° transforms the +Y axis into the -X axis.
    ' An additional positive rotation of 90° transforms the -X axis into the -Y axis.
    ' An additional positive rotation of 90° transforms the -Y axis into the +X axis (back where we started).
    ' =======================================================================================================
    With MatrixRotationZ
        .rc11 = sngCosine
        .rc21 = sngSine
        .rc12 = -sngSine
        .rc22 = sngCosine
    End With
    
    ' Very important note about this matrix
    ' =====================================
    ' If you see other programmers placing their Sines and Cosines in different positions (like the columns
    ' and rows have been swapped over - ie. Transposed) then this probably means that they have coded all
    ' of their algorithims to a different "notation standard". This subroutine follows the conventions used
    ' in the ledgendary bible "Computer Graphics Principles and Practice", Foley·vanDam·Feiner·Hughes which
    ' illustrates mathematical formulas using Column-Vector notation. Other books like "3D Math Primer for
    ' Graphics and Game Development", Fletcher Dunn·Ian Parberry, use Row-Vector notation. Both are correct,
    ' however it's important to know which standard you code to, because it affects the way in which you
    ' build your matrices and the order in which you should multiply them to obtain the correct result.
    '
    ' OpenGL uses Column Vectors (like this application).
    ' DirectX uses Row Vectors.

End Function

