Attribute VB_Name = "mDataStructures"
Option Explicit

' =========================================================================================
' 3D Computer Graphics for Visual Basic Programmers: Theory, Practice, Source Code and Fun!
' Version: 6.0 beta - Precision Edition
'
' by Peter Wilson
' Copyright © 2004 - Peter Wilson - All rights reserved.
' http://dev.midar.com/
' =========================================================================================


' =========================================================================
' This is a 4 dimensional vector, because it holds 4 values (X, Y, Z & W).
' Now you understand multi-dimensional vectors. See?, Vectors are not hard!
' =========================================================================
Public Type mdrVector4
    x As Single
    y As Single
    Z As Single
    w As Single         ' Named 'w' because we ran out of letters!
                        ' w is not often used (so you can optimise lots of code because of this.)
End Type


' =======================================
' A 4x4 Matrix - RC stands for RowColumn.
' =======================================
Public Type mdrMatrix4
    rc11 As Single: rc12 As Single: rc13 As Single: rc14 As Single
    rc21 As Single: rc22 As Single: rc23 As Single: rc24 As Single
    rc31 As Single: rc32 As Single: rc33 As Single: rc34 As Single
    rc41 As Single: rc42 As Single: rc43 As Single: rc44 As Single
End Type



' ============================================================
' Vertices are the simplest building blocks in 3D.
' ie.
'    A 2D triangle is made up from only 3 Vertices.
'    A 3D pyramid is made up from 5 Vertices.
'    A 3D cube is made up from 8 Vertices.
' Remember, a vertex is just a single point (or dot) in space.
' ============================================================
Public Type mdrVertex
    Pxyz As mdrVector4              '   Original Point - This is a permanent storage area.
    Txyz As mdrVector4              '   Transformed Point - This is a temporay storage area.
    Wxyz As mdrVector4              '   This is a temporay storage area.
    
    Brightness As Double            '   Any positive value between 0 and 1
    Clipped As Boolean              '   Polygon is partially visible and needs to be clipped.
End Type


Public Type mdr3DPart
    Caption As String                   '   Helicopter Blades, Landing Gear, Gun Turret, Leg, Head, Arm, etc. (Optional)
    Description As String               '   A Caption should always have a Description. (Optional)
    Selected As Boolean                 ' General purpose: Determins if the object is selected or not.
        
    Vertices() As mdrVertex             '   The original vertices that make up the object (these never changed once defined)
    Faces() As Variant                  '   Connect the dots [Vertices] together to form shapes.
    
    IdentityMatrix As mdrMatrix4        '   This holds the initial or default starting position for the polyhedron (rotation, size & position). (Optional)
End Type



' ======================================================================
' A 3D object is usually a collection of smaller objects (ie. Parts)
' ======================================================================
Public Type mdr3DObject
    ID              As String       ' General purpose: Reference number or string.
    Caption         As String       ' Helicopter, Tank, Space Ship, Monster, etc. (Optional)
    Description     As String       ' A Caption should always have a Description. (Optional)
    WorldPosition   As mdrVector4   ' Position of the Object in World Coordinates.
    Parts()         As mdr3DPart    ' This object is made up from Parts.
End Type


' ============================================
' This is our Virtual 3D Target Camera object.
' ============================================
Public Type mdr3DTargetCamera
    ID              As String       ' General purpose: Reference number or string.
    Class           As String       ' Class of Object
    Title           As String       '
    
    Visible         As Boolean      ' General purpose: Visible or Hidden from GUI.
    Caption         As String       ' Camera1, Director's Chair, Birds-eye View, etc. (Optional)
    Description     As String       ' A Caption should always have a Description.     (Optional)
    
    WorldPosition   As mdrVector4   ' VRP - Position of the Camera in World Coordinates.
    LookAtPoint     As mdrVector4   ' This is where the Camera is looking at in World Coordinates.
    VUP             As mdrVector4   ' Which way is UP?
    PRP             As mdrVector4   ' Projection Reference Point (PRP). Used for perspective distortion & stereopsis.
    
    Umin            As Single       '   The UV coordinate system coincides with the screen's XY coordinates.
    Umax            As Single       '       "
    Vmin            As Single       '       "
    Vmax            As Single       '       "
     
    ClipFar         As Single       ' Specified relative to VRP. Positive distance in the direction of VPN. This value is usually positive.
    ClipNear        As Single       ' Specified relative to VRP. Positive distance in the direction of VPN. This value is usually negative.
        
    ViewMatrix      As mdrMatrix4   ' View Matrix.
End Type


