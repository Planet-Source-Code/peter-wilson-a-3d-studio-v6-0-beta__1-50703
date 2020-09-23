VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Private Const mc_ModuleName As String = "Peters3DStudio.Application"


' This application class now owns a copy of the MDI form.
' We can now "recieve events" from the MDIForm (WithEvents keyword).
' ==================================================================
Private WithEvents m_frmMDI As frmMDI
Attribute m_frmMDI.VB_VarHelpID = -1

' This is where most of the drawing occurs.
' =========================================
Private WithEvents m_frmCanvas As frmCanvas
Attribute m_frmCanvas.VB_VarHelpID = -1

' This big-fat array holds ALL 3D objects in this aplication!
' (Arrays are much faster than VB Classes & Collections)
' ===========================================================
Private m_AllObjects()          As mdr3DObject


Private m_Camera                As mdr3DTargetCamera
Private m_matViewOrientation    As mdrMatrix4
Private m_matViewMapping        As mdrMatrix4
Private m_matViewPort           As mdrMatrix4


Private m_blnMouseCapture As Boolean

Private Sub DrawCrossHairs(p_PictureBox As PictureBox)

    ' Draws cross-hairs going through the origin of the 2D window.
    ' ============================================================
    p_PictureBox.DrawWidth = 1
    
    ' Draw Horizontal line (slightly darker to compensate for CRT monitors)
    p_PictureBox.Line (p_PictureBox.ScaleLeft, (p_PictureBox.ScaleHeight / 2))-(p_PictureBox.ScaleWidth, (p_PictureBox.ScaleHeight / 2)), RGB(160, 160, 160)
    
    ' Draw Vertical line
    p_PictureBox.Line ((p_PictureBox.ScaleWidth / 2), p_PictureBox.ScaleTop)-((p_PictureBox.ScaleWidth / 2), p_PictureBox.ScaleHeight), RGB(160, 160, 160)
    
End Sub

Private Sub RenderPipeline()

    On Error GoTo errTrap
    
    ' ======================================================
    ' Render Pipeline.
    ' ------------------------------------------------------
    ' 1) Database Traversal.
    ' 2) Model Transformation (move objects into the world).
    ' 3) Trivial accept/reject.
    ' 4) Lighting and Back-face culling.
    ' 5) View Transformation.
    ' 6) Clipping.
    ' 7) Divide by w. Map to 3D viewport.
    ' 8) Rasterization.
    ' 9) Display.
    ' ======================================================
    '
    '
    ' ======================================================
    ' Prerequisites
    ' ------------------------------------------------------
    ' Define the Camera
    ' ======================================================
        m_Camera.ID = "Camera1"
        m_Camera.Class = "Camera"
        m_Camera.Title = "Director's Chair"

        m_Camera.LookAtPoint.x = 0#
        m_Camera.LookAtPoint.y = 0#
        m_Camera.LookAtPoint.Z = 0#
        m_Camera.LookAtPoint.w = 1#
        
        m_Camera.VUP.x = 0#
        m_Camera.VUP.y = 1#
        m_Camera.VUP.Z = 0#
        m_Camera.VUP.w = 1#
        
        m_Camera.PRP.x = 0#
        m_Camera.PRP.y = 0#
        m_Camera.PRP.Z = 1#
        m_Camera.PRP.w = 1#
        
        m_Camera.Umin = -1#
        m_Camera.Umax = 1#
        m_Camera.Vmin = -1#
        m_Camera.Vmax = 1#
        
        m_Camera.ClipNear = -16
        m_Camera.ClipFar = -160
    
    
    Call DoSetDrawingParamaters(m_Camera)
    
    m_frmCanvas.pictCanvas(0).ScaleMode = vbPixels
    
    ' =======================================================
    ' Map canonical view volume into 3D viewport view volume.
    ' =======================================================
    With m_frmCanvas.pictCanvas(0)
        'm_matViewPort = Matrix_vv3dv(.ScaleLeft, .ScaleWidth, .ScaleTop, .ScaleHeight, -1, 0)
        m_matViewPort = Matrix_vv3dv(.ScaleLeft, .ScaleWidth, .ScaleHeight, .ScaleTop, -1, 0)
        m_frmCanvas.pictCanvas(0).BackColor = vbBlack
        m_frmCanvas.pictCanvas(0).Cls
    End With
    
    
    Call DrawCrossHairs(m_frmCanvas.pictCanvas(0))
    
    
    Dim intObject As Integer
    For intObject = LBound(m_AllObjects) To (UBound(m_AllObjects) - 1)
        With m_AllObjects(intObject)
        
            ' Draw each Part within the 3DObject.
            Dim lngPart As Long
            For lngPart = 0 To UBound(.Parts)
            
                Dim matOutput As mdrMatrix4
                Dim matWorld As mdrMatrix4
                
                matOutput = MatrixIdentity()
                matWorld = .Parts(lngPart).IdentityMatrix
                
                matOutput = MatrixMultiply(matOutput, matWorld)
                matOutput = MatrixMultiply(matOutput, m_matViewOrientation)
                matOutput = MatrixMultiply(matOutput, m_matViewMapping)
                
                
                ' For each vertex in the Part, recalculate.
                Dim blnClipMe As Boolean
                Dim sngW As Double
                
                Dim intJ As Integer
                For intJ = 0 To UBound(.Parts(lngPart).Vertices)
                                        
                    With .Parts(lngPart).Vertices(intJ)
                                            
                        ' Do some heavy number crunching here.
                        .Wxyz = MatrixMultiplyVector(matOutput, .Pxyz)
                        
                        .Clipped = False
                        sngW = .Wxyz.w
                        If sngW > 0 Then
                            ' W is positive.
                            If (.Wxyz.x < -sngW) Or (.Wxyz.x > sngW) Then .Clipped = True
                            If (.Wxyz.y < -sngW) Or (.Wxyz.y > sngW) Then .Clipped = True
                            If (.Wxyz.Z < -sngW) Or (.Wxyz.Z > 0) Then .Clipped = True
                        Else
                            ' W is negative
                            If (.Wxyz.x < -sngW) Or (.Wxyz.x > sngW) Then .Clipped = True
                            If (.Wxyz.y < -sngW) Or (.Wxyz.y > sngW) Then .Clipped = True
                            If (.Wxyz.Z < -sngW) Or (.Wxyz.Z > 0) Then .Clipped = True
                        End If
                        
                        If (.Clipped = False) Then
                            .Brightness = 1 - Abs(.Wxyz.Z)
                            If .Brightness > 1 Then .Brightness = 1
                        End If
                        
                        
                    End With
                Next intJ
SkipOver:
            Next lngPart
        
        
            On Error GoTo errTrap2
            For lngPart = 0 To UBound(.Parts)
                For intJ = 0 To UBound(.Parts(lngPart).Vertices)
                    With .Parts(lngPart).Vertices(intJ)
                        .Txyz = MatrixMultiplyVector(m_matViewPort, .Wxyz)
                    End With
                Next intJ
SkipOver2:
            Next lngPart
            

            ' Convert from 4D down to 3D
            ' ==========================
            Call Convert4Dto3D(m_AllObjects(intObject))


            ' ==============
            ' Draw 3D Object
            ' ==============
            Call Draw_SimpleWireframe(m_AllObjects(intObject))
            
            
        End With
    Next intObject
    
    Exit Sub
errTrap:
    Resume SkipOver
    
errTrap2:
    Resume SkipOver2
    
End Sub
Private Sub Draw_SimpleWireframe(withObject As mdr3DObject)
    
    ' Displays 3D usin Simple Wireframe

    On Error GoTo errTrap

    Dim intN As Long, intJ As Long, intK As Long
    Dim lngIndex0 As Long
    Dim vect0 As mdrVector4
    Dim lngIndex1 As Long
    Dim vect1 As mdrVector4
    Dim lngIndex2 As Long
    Dim vect2 As mdrVector4
    Dim blnClipped As Boolean
    Dim intClipCount As Integer

    
    With withObject

        ' Draw each object's polyhedra
        For intN = 0 To UBound(.Parts)
            m_frmCanvas.pictCanvas(0).DrawWidth = 1
            m_frmCanvas.pictCanvas(0).ForeColor = vbBlack
            
            ' Draw each polygon for the given polyhedra.
            For intJ = 0 To UBound(.Parts(intN).Faces)

                If .Parts(intN).Selected = True Then
                    m_frmCanvas.pictCanvas(0).ForeColor = vbWhite
                End If

                With .Parts(intN)
                
                
                    lngIndex0 = .Faces(intJ)(0)
                    lngIndex1 = .Faces(intJ)(1)
                    lngIndex2 = .Faces(intJ)(2)
                    
                    ' Enough faces for a Triangle?
                    If (UBound(.Faces(intJ)) > 1) Then
                    
                        ' Should the triangle be clipped?
                        blnClipped = False
                        blnClipped = blnClipped Or .Vertices(lngIndex0).Clipped
                        blnClipped = blnClipped Or .Vertices(lngIndex1).Clipped
                        blnClipped = blnClipped Or .Vertices(lngIndex2).Clipped
                        
                        intClipCount = 0
                        If .Vertices(lngIndex0).Clipped = True Then intClipCount = intClipCount + 1
                        If .Vertices(lngIndex1).Clipped = True Then intClipCount = intClipCount + 1
                        If .Vertices(lngIndex2).Clipped = True Then intClipCount = intClipCount + 1
                        
                        If blnClipped = True Then
                            vect0 = .Vertices(lngIndex0).Txyz
                            vect1 = .Vertices(lngIndex1).Txyz
                            vect2 = .Vertices(lngIndex2).Txyz
                            If intClipCount = 3 Then
'                                Call DrawFlatTriangle(m_frmCanvas.pictCanvas(0), vect0.X, vect0.Y, vect1.X, vect1.Y, vect2.X, vect2.Y, RGB(255, 0, 0), 0, vbSolid, vbFSTransparent)
                            Else
                                Call DrawFlatTriangle(m_frmCanvas.pictCanvas(0), vect0.x, vect0.y, vect1.x, vect1.y, vect2.x, vect2.y, RGB(0, 255, 255), 0, vbSolid, vbFSTransparent)
                            End If
                            
                        Else
                            vect0 = .Vertices(lngIndex0).Txyz
                            vect1 = .Vertices(lngIndex1).Txyz
                            vect2 = .Vertices(lngIndex2).Txyz
                            
                            m_frmCanvas.pictCanvas(0).ForeColor = RGB(0, Abs(255 * .Vertices(lngIndex0).Brightness), Abs(255 * .Vertices(lngIndex0).Brightness))

                            Call DrawFlatTriangle(m_frmCanvas.pictCanvas(0), vect0.x, vect0.y, vect1.x, vect1.y, vect2.x, vect2.y, m_frmCanvas.pictCanvas(0).ForeColor, 0, vbSolid, vbFSTransparent)
                            
                        End If ' Clipped=False
                    End If ' Enough faces for a Triangle?
                End With
SkipOver1:
            Next intJ
        Next intN
    End With
        
    
    m_frmCanvas.pictCanvas(0).Refresh
    
    
    Exit Sub
errTrap:
    Select Case Err.Number
        Case 9 ' Subscript Error
            Resume SkipOver1

        Case Else
            MsgBox Err.Number & " - " & Err.Description, vbCritical
            ' Do nothing

    End Select
    
End Sub



Private Sub Convert4Dto3D(p_3DObject As mdr3DObject)

    On Error GoTo errTrap
    
    Dim lngN As Long, lngJ As Long
    
    With p_3DObject
    
        ' For each object, calculate it's Parts
        For lngN = 0 To UBound(.Parts)
        
            ' For each vertex in the polyhedra, recalculate screen coordinates.
            ' This is a quick and nasty method of viewing 3D objects.  It is VERY crude but effective.
            For lngJ = 0 To UBound(.Parts(lngN).Vertices)
                With .Parts(lngN).Vertices(lngJ)
                
                    ' Ignore invalid co-ordinates
                    If (.Txyz.w = 0) = False Then
                        ' Convert from 4D down to 3D
                        ' ==========================
                        ' Apply the Perspective transformation by converting 4 dimensions down to
                        ' 3 dimensions, by dividing the x, y & z co-ordinates with w. Don't forget the w
                        ' value was previous calculated using the 'MatrixPerspective' routine.
                        ' I personally like to call this part of the code, 'dimensionally downshifting'.
                        .Txyz.x = .Txyz.x / .Txyz.w
                        .Txyz.y = .Txyz.y / .Txyz.w
                        .Txyz.Z = .Txyz.Z / .Txyz.w
                        
                    End If
                End With
SkipPointA:
            Next lngJ ' .Parts(lngN).Vertices
        Next lngN ' .Parts
    End With ' p_3DObject
    
    
    Exit Sub
errTrap:
    Select Case Err.Number
        Case 9 ' Subscript out of range
            Resume SkipPointA
            
        Case Else
            MsgBox Err.Number & " - " & Err.Description, vbExclamation
            
    End Select
    
End Sub

Private Sub DoSetDrawingParamaters(p_Camera As mdr3DTargetCamera)
    
    ' ======================================================================================
    '                       *** Synthetic [Virtual] Camera Values ***
    ' The four basic Viewing Parameters that define the View Reference Coordinates (VRC).
    ' They are all specified in World Coordinates (WC) except PRP; this is specified in VRC.
    '                       *** Synthetic [Virtual] Camera Values ***
    '
    ' ======================================================================================
    
    Dim vectVRP As mdrVector4   ' View Reference Point (VRP) - The world position of the virtual camera AND the virtual film!
    Dim vectVPN As mdrVector4   ' View Plane Normal (VPN) - The direction that the virtual camera is pointing "away from"!
    Dim vectPRP As mdrVector4   ' Projection Reference Point (PRP), also known as Centre Of Projection (COP) - This is the distance between the virtual camera's film, and the pin-hole lens of the virtual camera.
    Dim vectVUP As mdrVector4   ' View UP direction (VUP) - Which way is up? This is used for tilting (or not tilting) the camera.
    
    
    ' Define the View Reference Point (VRP)
    ' This is defined in the World Coordinate (WC) system.
    vectVRP = p_Camera.WorldPosition
    
    
    ' Subtract the Camera's world position (VRP) from the 'LookingAt' point to give us the View Plane Normal (VPN).
    ' VPN means different things to different 3D packages, ie. PHIGS and OpenGL do not agree on this one.
    ' In this application, the VPN points in the opposite direction that the camera is facing!! I said, Opposite!
    vectVPN = VectorSubtract(vectVRP, p_Camera.LookAtPoint)
    If (vectVPN.x = 0) And (vectVPN.y = 0) And (vectVPN.Z = 0) Then
        vectVPN.x = 0# ' Do not allow VPN to be all zero's (shouldn't happen anyway, but still check)
        vectVPN.y = 0#
        vectVPN.Z = 1#
    End If
    
    
    ' PRP is specified in the View Reference Coordinate system (and NOT the world coordinate system)
    vectPRP = p_Camera.PRP
'    vectPRP.x = 0#
'    vectPRP.y = 0#
'    vectPRP.z = 1#  ' << Change this value for perspective distortion (any positive value).
'    vectPRP.w = 1#
    
    
    ' The VUP vector is usually x=0,y=1,z=0. This is used to tilt the camera.
    vectVUP.x = 0#
    vectVUP.y = 1#
    vectVUP.Z = 0#
    vectVUP.w = 1#
    
    ' ============================================================================
    ' View Orientation.
    ' ============================================================================
    m_matViewOrientation = MatrixViewOrientation(vectVPN, vectVUP, vectVRP)
    
    
    ' ============================================================================
    ' Map projection view volume into canonical view volume suitable for clipping.
    ' ============================================================================
    m_matViewMapping = MatrixViewMapping_Per(p_Camera)
    
End Sub
Public Sub ShowApplication()

    ' ===================================
    ' Set and Load the form if necessary.
    ' ===================================
    If m_frmMDI Is Nothing Then
        Set m_frmMDI = New frmMDI
        Load m_frmMDI ' (Optional)
        
        ' Load Resources
        Call LoadResourceStrings(m_frmMDI)
    End If
    
    
    ' ========================================
    ' Show the form and bring it to the front.
    ' ========================================
    m_frmMDI.Show
    m_frmMDI.WindowState = vbMaximized
    
    
    ' =======================
    ' Call the Reset routine.
    ' =======================
    Call m_frmMDI_OnReset
    
End Sub

Public Function TearDown(Optional ForceTearDown As Boolean = False) As Boolean

    ' Assume Success (because Forms & Classes may not be loaded)
    TearDown = True
        
    
    ' =====================
    ' Attempt to close form
    ' =====================
    If Not (m_frmCanvas Is Nothing) Then
        ' Attempt to unload the form (which "may" trigger a save event if form is dirty)
        If m_frmCanvas.IsFormLoaded = True Then
            ' Unload controls/windows (but not code)
            Unload m_frmCanvas
        End If
        If m_frmCanvas.IsFormLoaded = False Then
            Set m_frmCanvas = Nothing ' Destroy [Code] Resources
        Else
            ' User must have clicked the "Cancel" button when prompted to save changes,
            ' or an error occured when they tried to save their changes while the application was closing.
            TearDown = False
        End If
    End If
    
        
End Function


Private Sub m_frmCanvas_OnMouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = vbLeftButton Then
        m_Camera.WorldPosition.x = -x + (m_frmCanvas.pictCanvas(0).ScaleWidth / 2)
        m_Camera.WorldPosition.y = y - (m_frmCanvas.pictCanvas(0).ScaleHeight / 2)
    ElseIf Button = vbRightButton Then
        m_Camera.WorldPosition.Z = x - (m_frmCanvas.pictCanvas(0).ScaleWidth / 2)
    
    End If
    
    m_frmCanvas.Caption = "x: " & m_Camera.WorldPosition.x & " y:" & m_Camera.WorldPosition.y & " z:" & m_Camera.WorldPosition.Z
    
    Call RenderPipeline
    
End Sub

Private Sub m_frmCanvas_OnRefresh()

    Screen.MousePointer = vbHourglass

    Call RenderPipeline
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub m_frmCanvas_OnResize()

    ' Reset mini-windows.
    With m_frmCanvas
        .pictCanvas(0).ScaleMode = vbPixels
        .pictCanvas(0).Move .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight
        .pictCanvas(0).Visible = True
    End With
    
End Sub

Private Sub m_frmMDI_OnImportDirectXDataFile()

    On Error GoTo errTrap
    
    Dim strFilePath As String
    Dim strFileTitle As String
    
       ' ================================
    ' Prompt user for location of file
    ' ================================
    With m_frmMDI.CommonDialog1
        .CancelError = True
        .Flags = cdlOFNHideReadOnly
        .DialogTitle = LoadResString(120)
        .FileName = ""
        .Filter = LoadResString(121)
        .ShowOpen
        
        strFilePath = .FileName
        strFileTitle = .FileTitle
        If strFilePath = "" Then Exit Sub
    End With
    
'    Call MsgBox("This application uses a Right-Handed coordinate system, but DirectX uses a Left-Hand coordinate system." & vbCrLf & vbCrLf & _
                "The DirectX data will now be loaded and converted to the Right-Hand system.", vbInformation)
    
    Screen.MousePointer = vbHourglass
    
    ' Reserve space for a new object
    Dim intObjectCount As Integer
        intObjectCount = UBound(m_AllObjects)
        intObjectCount = intObjectCount + 1
        ReDim Preserve m_AllObjects(intObjectCount)
    
    ' Load the DirectX file.
    m_AllObjects(intObjectCount - 1) = LoadXFile(strFilePath)
    
    ' Reset some basic properties of the object.
    m_AllObjects(intObjectCount - 1).ID = "Index=" & (intObjectCount - 1) & "&FileTitle=" & strFileTitle
    m_AllObjects(intObjectCount - 1).Caption = strFileTitle
    m_AllObjects(intObjectCount - 1).Description = strFilePath
        
    
    ' ===================
    ' Render all objects.
    ' ===================
    Call RenderPipeline
    
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
errTrap:
    Err.Source = mc_ModuleName & ".OnImportDirectXDataFile"
    Call LogAnError(Err, True)
    
End Sub

Private Sub m_frmMDI_OnReset()

    Dim intAnswer As Integer
    
    If (m_frmCanvas Is Nothing) = False Then
        intAnswer = MsgBox("Do you really want to reset?", vbYesNo + vbDefaultButton2 + vbQuestion, "Reset?")
        If intAnswer <> vbYes Then Exit Sub
    End If
    
    ' =============================
    ' Reset everything to defaults.
    ' =============================
    
    ' ======================================
    ' Teardown (ie. destroy) existing forms.
    ' ======================================
    If TearDown(False) = True Then
        
        ' Slow the reset process just a "tiny" bit, otherwise it might look
        ' like nothing has happened, and the user might get confused.
        DoEvents
        Call Sleep(500)
        
        ' Create the Canvas window.
        Set m_frmCanvas = New frmCanvas
        Load m_frmCanvas
        Call LoadResourceStrings(m_frmCanvas)
        m_frmCanvas.Show
        m_frmCanvas.Move 0, 0
        
        
        m_frmCanvas.Height = 6795
        m_frmCanvas.Width = m_frmCanvas.Height
        
        ' Enable the Import menu.
        m_frmMDI.mnuFileItem(2).Enabled = True
                
        ' Reset Memory
        ReDim m_AllObjects(0) As mdr3DObject
        
        ' Reset Camera Position.
        m_Camera.WorldPosition.x = 0#
        m_Camera.WorldPosition.y = 0#
        m_Camera.WorldPosition.Z = 128#
        m_Camera.WorldPosition.w = 1#
        
    End If
    
End Sub

