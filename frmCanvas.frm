VERSION 5.00
Begin VB.Form frmCanvas 
   AutoRedraw      =   -1  'True
   Caption         =   "300"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   Icon            =   "frmCanvas.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3780
   ScaleWidth      =   4635
   Begin VB.PictureBox pictCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2085
      Index           =   0
      Left            =   0
      ScaleHeight     =   2085
      ScaleWidth      =   2085
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2085
   End
End
Attribute VB_Name = "frmCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Define the name of this class/module for error-trap reporting (optional).
Private Const m_strModuleName As String = "frmCanvas"

' Save Code.
Private m_blnFormDirty As Boolean

' Events
Event OnResize()
Event OnRefresh()
Event OnMouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

' TearDown code
Private m_blnStillLoaded As Boolean

Private Sub Form_Load()

    ' Teardown Code
    m_blnStillLoaded = True
    
    ' Save Code
    m_blnFormDirty = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' TearDown code
    m_blnStillLoaded = False

End Sub

Public Property Get IsFormLoaded() As Boolean

    ' TearDown code
    IsFormLoaded = m_blnStillLoaded
    
End Property

Public Property Get IsFormDirty() As Boolean

    IsFormDirty = m_blnFormDirty

End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        
    Dim intA As Integer
    
    ' Check if the Microsoft Windows Task Manager is closing the application.
    If g_intUnloadMode = vbAppTaskManager Then
        ' Discard changes and don't allow user to interupt TaskManager shutdown
        ' (Basically I hate it when I'm trying to force a shut down and it does not happen!)
        m_blnFormDirty = False
    End If
    
    ' Only prompt user if form is dirty
    If Me.IsFormDirty = True Then
        Me.ZOrder
        intA = MsgBox("Save changes?", vbQuestion + vbYesNoCancel + vbDefaultButton1, Me.Caption)
        Select Case intA
            Case vbYes
'                If ApplyChanges = False Then Cancel = True
                
            Case vbNo
                ' Do nothing - form will unload
                
            Case vbCancel
                ' Cancel form closure
                Cancel = True
                
        End Select
    End If
    
End Sub

Private Sub Form_Resize()
    
    RaiseEvent OnResize
    
End Sub

Private Sub pictCanvas_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Send the event back to our parent.
    RaiseEvent OnMouseDown(Index, Button, Shift, x, y)
    
End Sub


