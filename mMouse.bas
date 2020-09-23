Attribute VB_Name = "mMouse"
Option Explicit

' =========================================================================================
' 3D Computer Graphics for Visual Basic Programmers: Theory, Practice, Source Code and Fun!
' Version: 5.0 - Game Edition.
'
' by Peter Wilson
' Copyright © 2003 - Peter Wilson - All rights reserved.
' http://dev.midar.com/
' =========================================================================================


' Define the name of this class/module for error-trap reporting.
Private Const m_strModuleName As String = "mMouse"


' API Declarations used to GET & SET the position of the mouse.
Public Type POINTAPI
        x As Long
        y As Long
End Type
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Attribute GetCursorPos.VB_Description = "GetCursorPos reads the current position of the mouse cursor."
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Attribute SetCursorPos.VB_Description = "SetCursorPos sets the current position of the mouse cursor."

' The ShowCursor function displays or hides the cursor.
' (This function sets an internal display counter that determines whether the cursor should be displayed.
'  The cursor is displayed only if the display count is greater than or equal to 0.
'  If a mouse is installed, the initial display count is 0. If no mouse is installed, the display count is –1.)
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Attribute ShowCursor.VB_Description = "The ShowCursor function displays or hides the cursor."

