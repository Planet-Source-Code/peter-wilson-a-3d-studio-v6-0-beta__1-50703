Attribute VB_Name = "mRasterization"
Option Explicit

' =========================================================================================
' 3D Computer Graphics for Visual Basic Programmers: Theory, Practice, Source Code and Fun!
' Version: 6.0 beta - Precision Edition
'
' by Peter Wilson
' Copyright © 2004 - Peter Wilson - All rights reserved.
' http://dev.midar.com/
' =========================================================================================

' API Declarations used for drawing and filling "flat-shaded" triangles.
Public Type POINT_TYPE
  x As Long
  y As Long
End Type
Public Declare Function Polygon Lib "gdi32.dll" (ByVal hDC As Long, lpPoint As POINT_TYPE, ByVal nCount As Long) As Long

Private Function convert_Long2UShort(ULong As Long) As Integer

    ' This function converts a long integer to an unsigned integer (ie. a UShort)
    '
    ' All Visual Basic integers are "signed" integers, meaning they go from -32,768 to 32,767
    ' However our API routine needs an "unsigned" integer, meaning it goes from 0 to 65534.
    ' Both signed, and unsigned integers take up 16 bits. ie. 0000000000000000
    
    If ULong <= &H7FFF& Then
        convert_Long2UShort = ULong
    Else
        convert_Long2UShort = Not (&HFFFF& - ULong)
    End If
    
    
    ' Update debug & info. counters (conditionally compiled)
    #If g_blnCompiledDebugInfo = True Then
        g_lngSubtractionCount = g_lngSubtractionCount + 1
    #End If
    
End Function

Public Sub DrawFlatTriangle(withHandle As PictureBox, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double, lngDrawColour As Long, lngFillColour As Long, intDrawStyle As DrawStyleConstants, intFillStyle As FillStyleConstants)
    
    ' Do basic error checking.
    If withHandle Is Nothing Then Exit Sub
    If withHandle.HasDC = False Then Exit Sub
    
    Dim points(0 To 3) As POINT_TYPE
    
    points(0).x = x1: points(0).y = y1
    points(1).x = x2: points(1).y = y2
    points(2).x = x3: points(2).y = y3
    points(3) = points(0)
    
    ' Fill Options
    ' =============
    If lngFillColour = -1 Then ' Turn off fill.
        withHandle.FillStyle = vbFSTransparent
    Else
        ' Fill polygon with specified colour.
        withHandle.FillColor = lngFillColour
        withHandle.FillStyle = intFillStyle
    End If
    
    ' Draw Options
    ' ============
    If lngDrawColour = -1 Then ' Turn off edges
        withHandle.DrawStyle = vbInvisible
    Else
        withHandle.ForeColor = lngDrawColour
        withHandle.DrawStyle = intDrawStyle
    End If
    
    ' Call the API to Draw and Fill the polygon.
    Call Polygon(withHandle.hDC, points(0), 4)
    
End Sub



