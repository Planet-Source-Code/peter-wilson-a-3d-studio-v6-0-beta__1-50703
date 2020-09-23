VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H80000001&
   Caption         =   "102"
   ClientHeight    =   5745
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8040
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3780
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "200"
      Begin VB.Menu mnuFileItem 
         Caption         =   "201"
         Index           =   0
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "202"
         Enabled         =   0   'False
         Index           =   2
         Begin VB.Menu mnuImportItem 
            Caption         =   "203"
            Index           =   0
         End
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   98
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "204"
         Index           =   99
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Event OnReset()
Event OnImportDirectXDataFile()

Private Sub mnuFileItem_Click(Index As Integer)

    Select Case Index
        Case 0 ' Reset
            RaiseEvent OnReset
            
        Case 99 ' Exit Application
            Unload Me
            
    End Select
    
End Sub

Private Sub mnuImportItem_Click(Index As Integer)

    Select Case Index
        Case 0 ' Import DirectX File (*.x)
            RaiseEvent OnImportDirectXDataFile
            
    End Select
    
End Sub

