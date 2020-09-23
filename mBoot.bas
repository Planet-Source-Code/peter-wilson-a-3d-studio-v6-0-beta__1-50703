Attribute VB_Name = "mBoot"
Option Explicit

' =========================================================================================
' 3D Computer Graphics for Visual Basic Programmers: Theory, Practice, Source Code and Fun!
' Version: 6.0 beta - Precision Edition
'
' by Peter Wilson
' Copyright © 2004 - Peter Wilson - All rights reserved.
' http://dev.midar.com/
' =========================================================================================

' Variable Prefixes used in this program
' ======================================
'   g_      =   global variable. ie. this variable is available anywhere within the program.
'   m_      =   module-level variable
'   p_      =   parameters that appear in function/sub declarations. ie. Private Sub DoWork(p_strWithData As String)
'   bln     =   Boolean Data Type. True or False
'   int     =   Integer Data Type. -32,768 to 32,767
'   lng     =   Long Data Type. (Long integer) -2,147,483,648 to 2,147,483,647
'   sng     =   Single Data Type. (single-precision floating-point) -3.402823E38 to -1.401298E-45 for negative values; 1.401298E-45 to 3.402823E38 for positive values
'   dbl     =   Double Data Type. (double-precision floating-point) -1.79769313486232E308 to -4.94065645841247E-324 for negative values; 4.94065645841247E-324 to 1.79769313486232E308 for positive values
'   cur     =   Currency Data Type. (scaled integer) -922,337,203,685,477.5808 to 922,337,203,685,477.5807
'   obj     =   Generic Object; either an unknown type of Object, or a known Object type.
'   cls     =   Class Object.
'   frm     =   Form Object.
'   var     =   Variant Data Type.
'   Example:
'       g_clsApplication        =   Global Class named Application
'       g_curFramePerfCount     =   Global Currency Data Type called FramePerfCount
'       m_frm3DLibrary          =   Module Level variable, which holds a Form object called 3DLibrary.


' How to Open and Close this application.
' =======================================
' This mBoot module creates the main Application Class (g_clsApplication).
' From now on - the Application class calls the shots and is basically in chrage.
' You are probably used to just Loading a Form, or even an MDI form.  Well... this
' is a little bit different. The Application class (g_clsApplication) is now in
' charge of the form 'frmCanvas'. You don't have to code this way, it's just my personal preference.
Public g_clsApplication As Peters3DStudio.Application


' Used in Form_QueryUnload events
Public g_intUnloadMode As Integer


' These counters are used for debugging & testing only (conditionally compiled with the # hash symbol)
' Change the value of 'g_blnCompiledDebugInfo' in the Project Properties area, NOT in the code.
#If g_blnCompiledDebugInfo = True Then
    Public g_lngMultiplicationCount As Long
    Public g_lngDivisionCount As Long
    Public g_lngAdditionCount As Long
    Public g_lngSubtractionCount As Long
    Public g_lngSquareRootCount As Long
    Public g_lngFrameTickCount As Long
    Public g_curFramePerfCount As Currency          ' We're NOT using money, but rather VB's 64-bit data type.
    Public g_curPerformanceFrequency As Currency    ' We're NOT using money, but rather VB's 64-bit data type.
#End If

Public Sub Main()

    ' =============================================================
    ' Start Application Logging (to a file)
    ' Note: Application logging does not work in VB's Runtime mode,
    '       it only works for compiled EXE's.
    ' =============================================================
    Call App.StartLogging(App.Path, vbLogAuto)
       
    
    ' =================================
    ' This kick-starts our Application!
    ' =================================
    If g_clsApplication Is Nothing Then Set g_clsApplication = New Peters3DStudio.Application
    Call g_clsApplication.ShowApplication
    ' Note: If the g_clsApplication class didn't open a form, then this VB program
    '       would end right here!  This VB program will end automatically when
    '       the g_clsApplication class closes all of the forms that it has opened.
    '
    '       In VB.NET it's not enough to show a form (within the application class), you must
    '       also start a "message pump"; VB6 automatically does this for you. I only include
    '       this here, because it was driving me nuts and the VB.NET help file sux and blows
    '       all at the same time!
    
End Sub

