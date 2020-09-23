Attribute VB_Name = "mErrorTrapping"
Option Explicit


Public Sub LogAnError(p_objErr As ErrObject, Optional p_blnDisplayMsg As Boolean = True)
    
    Dim strForEventLog As String
    Dim strForDisplay As String
    Dim strForWeb As String
    
    Screen.MousePointer = vbDefault

    ' ================================================================================================
    ' Build a string suitable for sending via an Internet URL.
    ' ie. http://www.[domain_name_goes_here].com/onlinecontent.asp?Err=6&Desc=Overflow&Source=Project1
    ' ================================================================================================
    strForWeb = "Err=" & p_objErr.Number & "&Desc=" & p_objErr.Description & "&Src=" & p_objErr.Source
    
    
    ' ========================================================
    ' Build a suitable string for the Display and/or EventLog.
    ' ========================================================
    If App.LogMode = vbLogToNT Then
        strForDisplay = LoadResString(108) & vbNewLine & vbNewLine & _
                        LoadResString(111) & p_objErr.Number & vbNewLine & _
                        LoadResString(112) & p_objErr.Description & vbNewLine & _
                        LoadResString(113) & p_objErr.Source & vbNewLine & vbNewLine & _
                        LoadResString(110) & vbNewLine
        strForEventLog = vbNewLine & strForDisplay
        
    ElseIf App.LogMode = vbLogToFile Then
        strForDisplay = LoadResString(108) & vbNewLine & vbNewLine & _
                        LoadResString(111) & p_objErr.Number & vbNewLine & _
                        LoadResString(112) & p_objErr.Description & vbNewLine & _
                        LoadResString(113) & p_objErr.Source & vbNewLine & vbNewLine & _
                        LoadResString(109) & vbNewLine & _
                        " ' " & App.LogPath & " ' "
        strForEventLog = Format(Now, "yyyy-mm-dd hh:mm:ssampm") & ", " & Replace(strForDisplay, vbNewLine, ", ")
        
    Else
        strForDisplay = LoadResString(108) & vbNewLine & vbNewLine & _
                        LoadResString(111) & p_objErr.Number & vbNewLine & _
                        LoadResString(112) & p_objErr.Description & vbNewLine & _
                        LoadResString(113) & p_objErr.Source
        strForEventLog = strForDisplay ' << Doesn't matter since it shouldn't save anyway.
        
    End If
    
    
    ' ==================================
    ' Display error message to the user.
    ' ==================================
    If p_blnDisplayMsg = True Then
        Call MsgBox(strForDisplay, vbCritical + vbMsgBoxHelpButton, LoadResString(108), App.HelpFile, 0)
    End If
    
    
    ' ============================================================================================
    ' Log error message to file or EventLog (works only for compiled applications - ie. EXE files)
    ' ============================================================================================
    App.LogEvent strForEventLog, vbLogEventTypeError
    
End Sub


