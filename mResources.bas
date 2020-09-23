Attribute VB_Name = "mResources"
Option Explicit

Public Sub LoadResourceStrings(CurrentForm As Form)
    
    On Error Resume Next

    Dim ctl As Control
    Dim obj As Object
    Dim fnt As Object
    Dim sCtlType As String
    Dim intN As Integer
    
    For Each ctl In CurrentForm.Controls
        sCtlType = TypeName(ctl)
        If sCtlType = "Label" Then
            ctl.Caption = LoadResString(CInt(ctl.Caption))
        ElseIf sCtlType = "Menu" Then
            ctl.Caption = LoadResString(CInt(ctl.Caption))
        ElseIf sCtlType = "ListView" Then
            For Each obj In ctl.ColumnHeaders
                obj.Text = LoadResString(CInt(obj.Text))
            Next
        ElseIf sCtlType = "SSTab" Then
            For intN = 0 To ctl.Tabs - 1
                ctl.TabCaption(intN) = LoadResString(CInt(ctl.TabCaption(intN)))
            Next intN
        ElseIf sCtlType = "CommandButton" Then
            ctl.Caption = LoadResString(CInt(ctl.Caption))
        ElseIf sCtlType = "Frame" Then
            ctl.Caption = LoadResString(CInt(ctl.Caption))
        ElseIf sCtlType = "TextBox" Then
            ctl.Text = LoadResString(CInt(ctl.Text))
        ElseIf sCtlType = "Image" Then
            ctl.ToolTip = LoadResString(CInt(ctl.ToolTip))
            ctl.Icon = LoadResPicture(CInt(ctl.Tag), vbResIcon)
            ctl.Picture = LoadResPicture(CInt(ctl.Tag), vbResIcon)
        End If
    Next
    
    CurrentForm.Caption = LoadResString(CInt(CurrentForm.Caption))

End Sub


