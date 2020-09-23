Attribute VB_Name = "mMaths"
Option Explicit

' For this application we only need a very-low precision value for PI; ie. The "Single" data type.
' ================================================================================================
Public Const m_sngPI As Single = 3.141594!
Public Const m_sngPIDivideBy180 As Single = 0.0174533!
Public Const m_sng180DivideByPI As Single = 57.29578!


' This is a 4 dimensional vector, because it holds 4 values (X, Y, Z & W).
' Now you understand multi-dimensional vectors. See?, Vectors are not hard!
' =========================================================================
Public Type mdrVector4
    X As Single
    Y As Single
    Z As Single
    W As Single         ' Named 'W' because we ran out of letters!
    DotColour As Long   ' << This isn't really the correct spot for this, but it was a last minute change. It will be moved to a better location in future versions.
End Type


' A 4x4 Matrix - RC stands for RowColumn
' ======================================
Public Type mdrMATRIX4
    rc11 As Single: rc12 As Single: rc13 As Single: rc14 As Single
    rc21 As Single: rc22 As Single: rc23 As Single: rc24 As Single
    rc31 As Single: rc32 As Single: rc33 As Single: rc34 As Single
    rc41 As Single: rc42 As Single: rc43 As Single: rc44 As Single
End Type

Public Function ConvertDeg2Rad(Degress As Single) As Single

    ' Converts Degrees to Radians
    ConvertDeg2Rad = Degress * (m_sngPIDivideBy180)
    
End Function



Public Function DotProduct(VectorU As mdrVector4, VectorV As mdrVector4) As Single

    ' Determines the dot-product of two 4D vectors.
    DotProduct = (VectorU.X * VectorV.X) + (VectorU.Y * VectorV.Y) + (VectorU.Z * VectorV.Z) + (VectorU.W * VectorV.W)
    
End Function

Public Function MatrixViewOrientation(vectVPN As mdrVector4, vectVUP As mdrVector4, vectVRP As mdrVector4) As mdrMATRIX4
    
    ' =====================================================
    ' Rotate VRC such that the:
    '   * n axis becomes the z axis,
    '   * u axis becomes the x axis and
    '   * v axis becomes the y axis.
    ' =====================================================
    
    Dim matRotateVRC As mdrMATRIX4
    Dim matTranslateVRP As mdrMATRIX4
    
    Dim vectN As mdrVector4
    Dim vectU As mdrVector4
    Dim vectV As mdrVector4
    
    
    '         VPN
    ' n = ¯¯¯¯¯¯¯¯¯¯¯
    '       | VPN |
    vectN = VectorNormalize(vectVPN)
    
    
    '       VUP x n
    ' u = ¯¯¯¯¯¯¯¯¯¯¯¯¯
    '     | VUP x n |
    vectU = CrossProduct(vectVUP, vectN)
    vectU = VectorNormalize(vectU)
    
    
    ' v = n x u
    vectV = CrossProduct(vectN, vectU)
    
    
    ' Define the Rotate matrix such that the n-axis (VPN) becomes the z-axis,
    ' the u-axis becomes the x-axis and the v-axis becomes the y-axis.
    matRotateVRC = MatrixIdentity()
    With matRotateVRC
        .rc11 = vectU.X: .rc12 = vectU.Y: .rc13 = vectU.Z
        .rc21 = vectV.X: .rc22 = vectV.Y: .rc23 = vectV.Z
        .rc31 = vectN.X: .rc32 = vectN.Y: .rc33 = vectN.Z
    End With
    
    
    ' Define a Translation matrix to transform the VRP to the origin.
    matTranslateVRP = MatrixTranslation(-vectVRP.X, -vectVRP.Y, -vectVRP.Z)
    
    
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

    ' Subtracts vector 2 away from vector 1.
    With VectorSubtract
        .X = V1.X - v2.X
        .Y = V1.Y - v2.Y
        .Z = V1.Z - v2.Z
    End With
    
End Function

Public Function MatrixTranslation(OffsetX As Single, OffsetY As Single, OffsetZ As Single) As mdrMATRIX4
    
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

Public Function MatrixIdentity() As mdrMATRIX4

    ' The identity matrix is used as the starting point for matrices
    ' that will modify vertex values to create rotations, translations,
    ' and any other transformations that can be represented by a 4×4 matrix
    '
    ' Notice that...
    '   * the 1's go diagonally down?
    '   * rc stands for Row Column. Therefore, rc12 means Row1, Column 2.
    
    With MatrixIdentity
        .rc11 = 1: .rc12 = 0: .rc13 = 0: .rc14 = 0
        .rc21 = 0: .rc22 = 1: .rc23 = 0: .rc24 = 0
        .rc31 = 0: .rc32 = 0: .rc33 = 1: .rc34 = 0
        .rc41 = 0: .rc42 = 0: .rc43 = 0: .rc44 = 1
    End With
    
End Function

Public Function MatrixMultiply(m1 As mdrMATRIX4, m2 As mdrMATRIX4) As mdrMATRIX4
    
    ' Re-declare m1 & m2
    Dim m1b As mdrMATRIX4
    Dim m2b As mdrMATRIX4
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
    ' Note: This was only a 3x3 matrix show... however this routine is bigger again, using a 4x4. I just wanted to keep the example short.
    
    
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

Public Function MatrixMultiplyVector(m1 As mdrMATRIX4, V1 As mdrVector4) As mdrVector4
        
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
    ' conventions, and it's quite easy to get them mixed up with each other.
    
    
    
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
        .X = (m1.rc11 * V1.X) + (m1.rc12 * V1.Y) + (m1.rc13 * V1.Z) + (m1.rc14 * V1.W)
        .Y = (m1.rc21 * V1.X) + (m1.rc22 * V1.Y) + (m1.rc23 * V1.Z) + (m1.rc24 * V1.W)
        .Z = (m1.rc31 * V1.X) + (m1.rc32 * V1.Y) + (m1.rc33 * V1.Z) + (m1.rc34 * V1.W)
        .W = (m1.rc41 * V1.X) + (m1.rc42 * V1.Y) + (m1.rc43 * V1.Z) + (m1.rc44 * V1.W)
    End With
    
End Function

Public Function VectorNormalize(v As mdrVector4) As mdrVector4

    ' Returns the normalized version of a 3-D vector.
    
    Dim sngLength As Single
    
    sngLength = VectorLength(v)
    If sngLength = 0 Then sngLength = 1
    
    With VectorNormalize
        .X = v.X / sngLength
        .Y = v.Y / sngLength
        .Z = v.Z / sngLength
    End With
    
End Function

Public Function VectorLength(v As mdrVector4) As Single

    ' Returns the length of a Vector.
    '
    ' In Mathematic books, the "length of a vector" is often written with two verticle bars on either
    ' side, like this:  ||v||
    ' It took me ages to figure this out! Nobody explained it, they just assumed I knew it!
    '
    ' The length of a vector is from the origin (0,0,0) to x,y,z
    ' Do you remember high schools maths, Pythagoras theorem?  c^2 = a^2 + b^2
    '   "In a right-angled triangle, the area of the square of the hypotenuse (the longest side)
    '    is equal to the sum of the areas of the squares drawn on the other two sides."
    
    VectorLength = Sqr((v.X ^ 2) + (v.Y ^ 2) + (v.Z ^ 2))
    
End Function

Public Function CrossProduct(vectV As mdrVector4, VectW As mdrVector4) As mdrVector4

    ' Determines the cross-product of two 3-D vectors (V and W).
    ' The cross-product is used to find a vector that is perpendicular
    ' to the plane defined by VectV and VectW.
    
    With CrossProduct
        .X = (vectV.Y * VectW.Z) - (vectV.Z * VectW.Y)
        .Y = (vectV.Z * VectW.X) - (vectV.X * VectW.Z)
        .Z = (vectV.X * VectW.Y) - (vectV.Y * VectW.X)
    End With
    
End Function

