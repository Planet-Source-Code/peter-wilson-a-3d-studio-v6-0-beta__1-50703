Attribute VB_Name = "mDirectXParser"
Option Explicit

' This module is a "Recursive Descent Parser" for parsing DirectX text files (*.x)
' Copyright © 2003-2004 Peter Wilson, Canberra ACT. See: http://dev.midar.com/
'
'   Recursive Descent Parser's are the kinds of things you learn when you go to University, or
'   do special courses in computer programming. These kinds of parsers are very powerful and are
'   almost always used to parse complex files (like computer code), or DirectX .X files ;-)
'   Many types of Compilers (C or Visual Basic) use parsers to make sure the user has typed in the
'   code the correct way without Syntax Errors.
'
'   I wrote this file parser because I did not want to call a single DirectX function - PURE VB!
'   DirectX actually contains helper routines to load .X files, however I didn't want to use them.
'
'   Here's how it works...
'   ======================
'   This parser works by matching patterns it finds in the data file, to VB Functions that then
'   process the pattern matched.
'
'   Step 1) Look for "Mesh" token using the following "Regular Expression":
'
'           ^Mesh .{1,}
'
'   Step 2) TokenFoundInFile = "Mesh"   --->>>  Run Visual Basic Subroutine called Parse_Mesh()
'
'   ie. If the GetNextToken() function finds the word: "mesh" then we call the Parse_Mesh() function.
'       The Parse_Mesh() function will then attempt to process whatever it finds (including other
'       tokens).
'       Note: The Parse_Frame() function, has the ability to process sub-Frames, thus this
'       function is said to be recursive - it calls itself!
'
'       This is the basic idea behind a parser. Get a token, match it, process it.
'
'   Limitations
'   ===========
'   This parser does NOT parse the whole DirectX file, only parts of it. Materials, animations, etc. are
'   ignored. Only one Header template per file allowed. ** Text format only **. Nested Frames are supported.
'   For multiple Meshes and Frames, they must be contained within a top level Frame template.
'
'   IMPORTANT
'   =========
'   This application uses Right-Handed Coordinate System, but DirectX uses a Left-Handed System.
'   This application will flip the z-axis sign, Also see: Parse_Mesh() function.
'
'   Cool Features
'   =============
'   Use the Conv3ds.exe program supplied with the DirectX SDK to convert 3DS files (how cool is that?!)
'   into .x format, then import them into this VB program and have fun!
'
'       Example: c:\conv3ds.exe -TMNctx myfile.3ds
'
'   If your 3D object looks screwed up after conversion, send me an e-mail
'   with the details and I'll try to help, or visit my web site's FAQ at http://dev.midar.com/
'
'   support@midar.com


' Used for error trapping routines
Private Const mc_ModuleName As String = "Peters3DStudio.mDirectXParser"


' The working text contains the raw file data.
Private m_strWorkingText As String

' DirectX .x Files are broken up by line-feed characters.
' This parser treats each line (or row) as a Token.
' Generally, there is a separate Function for each Token we will process.
' This variant array holds all of the tokens available. This includes
' reserved words, template names, brackets, data values, etc.
Private m_Tokens As Variant
Private m_lngCurrentToken As Long
Private m_strOldTokenValue As String        '   Old Token
Private m_strCurrentTokenValue As String    '   Current Token


' This is a temporary 3D object that the data will be imported into.
Private m_obj3DObject As mdr3DObject
Private m_lngPartCount As Long    '   each frame
Private m_lngVerticeCount As Long       '   each mesh
Private m_lngFaceCount As Long          '   each face

Public Function LoadXFile(FilePath As String) As mdr3DObject

    Dim strFullFile As String
    Dim strLineOfData As String
    
    
    ' ========================================================================
    ' Open the file and read the whole file into memory (ie. into strFullFile)
    ' ========================================================================
    Open FilePath For Input Access Read Lock Read As #1
        
        ' Read the WHOLE file into memory (ie. strFullFile)
        strFullFile = ""
        Do While EOF(1) = False ' Do this Loop while the End-Of-File is false (or until the End-Of-File = True)
        
            ' Read the next Row [Line] of data from the open file #1
            ' Note: This process reads a single line of data right up until the
            '       Carriage-Return/Line-Feed pair, but dosn't include the Carriage-Return/Line-Feed pair!
            Line Input #1, strLineOfData
    
            ' Note: This concatenation can be S-L-O-W for large files.
            '       Search the net for: "fast string concatenation" (if needed)
            ' Note: We're adding a Line-Feed character (ASCII Value 10) that was stripped in the above step.
            '       This is because DirectX text files use this to delimit different parts of the file.
            strFullFile = strFullFile & strLineOfData & vbLf
                        
        Loop
    Close #1 ' Close the file, we'll work from memory now (ie. strFullFile)
    
    
    ' ===================================================================
    ' Call the "Recursive Descent Parser" to parse this DirecX text file.
    ' ===================================================================
    Call CreateTokens(strFullFile, LoadXFile)
    
End Function

Private Sub RemoveEmptyFrames()

    ' When you convert 3DS files into .X file format using the following...
    '
    '       c:\conv3ds.exe -TMNctx myfile.3ds
    '
    ' ...then there is a likelyhood that some some .X files will contain an empty top-level Frame template,
    ' usually the more complex 3D objects that are made up from many different parts, and sub-parts, etc.
    ' This function removes any empty top-level Frame templates that have been converted to Part's.
    ' NOTE: DO NOT leave out the -T option (on conv3ds.exe) when converting .X files.
    '       Not all 3DS files require this option, however this parser has been optomised under the
    '       assumption that the -T option is present (ie. Include Top-Level Frame)
    
End Sub


Private Function TrimWhiteSpace(ByVal strString As String) As String

    ' This function is similar to VB's TRIM operator, except this one
    ' trims both TAB and SPACE characters (ie. Whitespace)
    
    Dim strChar As String
    Dim intN As Integer
    
    TrimWhiteSpace = Trim(strString)
    If TrimWhiteSpace = "" Then Exit Function
    
    ' Trim leading TAB or SPACE characters (ASCII values 9 and 32)
    strChar = Left(TrimWhiteSpace, 1)
    If strChar <> "" Then
        While (Asc(strChar) = 9) Or (Asc(strChar) = 32)
            TrimWhiteSpace = Mid(TrimWhiteSpace, 2)
            strChar = Left(TrimWhiteSpace, 1)
        Wend
    End If
    
    ' Trim trailing TAB or SPACE characters (ASCII values 9 and 32)
    strChar = Right(TrimWhiteSpace, 1)
    If strChar <> "" Then
        While (Asc(strChar) = 9) Or (Asc(strChar) = 32)
            TrimWhiteSpace = Left(TrimWhiteSpace, Len(TrimWhiteSpace) - 1)
            strChar = Right(TrimWhiteSpace, 1)
        Wend
    End If
    
    
End Function

Private Sub CreateTokens(ByVal WithWorkingText As String, obj3DObject As mdr3DObject)
    
    On Error GoTo errTrap
    
    Dim blnMatchFound As Boolean
    Dim blnFrameDone As Boolean
    Dim blnMeshDone As Boolean
    
    ' Set local variables
    m_strWorkingText = WithWorkingText
    m_obj3DObject = obj3DObject
    
    ' Remove any Carriage Returns (shouldn't be any, but check anyway)
    m_strWorkingText = Replace(m_strWorkingText, Chr(13), "")
    
    ' Split working text up into tokens (as separated by a line-feed ASCII value 10)
    m_Tokens = Split(m_strWorkingText, Chr(10))
    
    ' Reset Current Token indicator
    m_lngCurrentToken = -1
    
    ' Reset Part count
    m_lngPartCount = -1
    
    ' ==================================================================================
    ' Kick start the token process
    ' NOTE: This routine does not process ALL types of DirectX .X files, only files that
    '       conform to output produced using the following line command:
    '       conv3ds.exe -TMNctx myfile.3ds
    ' ==================================================================================
    Call GetNextToken
    Do
        blnMatchFound = False
        blnFrameDone = False
        blnMatchFound = False
        
        ' Search for Top-Level Tokens in the file.
        If MatchToken("^xof.{1,}$") = True Then
            
            ' Go and process the rest of the data matching the token pattern: ^xof.{1,}$
            blnMatchFound = Parse_FileHeader()
        
        ElseIf MatchToken("^Material {1,}") = True Then
            
            blnMatchFound = Parse_Material
        
        ElseIf MatchToken("^Template\b.{1,}{") = True Then
            
            blnMatchFound = Parse_Template
            
        ElseIf MatchToken("^Header {$") = True Then
            
            blnMatchFound = Parse_Header
        
        ElseIf MatchToken("^Frame .{1,}") = True Then   ' Note: This is a TOP-LEVEL frame, once it's finished there is no more to process.
            
            blnFrameDone = Parse_Frame
            'blnMatchFound = Parse_Frame
            
        ElseIf MatchToken("^Mesh .{1,}") = True Then    ' Note: This is a TOP-LEVEL frame, once it's finished there is no more to process.
        
            m_lngPartCount = m_lngPartCount + 1
            ReDim m_obj3DObject.Parts(m_lngPartCount)
                  m_obj3DObject.Parts(m_lngPartCount).IdentityMatrix = MatrixIdentity
            blnMeshDone = Parse_Mesh()
            
        Else
            ' This parser is designed to look for the above top-level "tokens" and if it doesn't find any
            ' of them by the time it gets to MESH, then the DirectX .x File is deemed to be in an
            ' unrecognized format (at least by this program anyway!)
            Err.Raise vbObjectError + 1001, "CreateTokens2", "Unrecognized file format. There was no valid Template, Header, Frame or Mesh object to process."
        End If
        
    Loop Until (blnFrameDone = True) Or (blnMeshDone = True) Or (blnMatchFound = False) Or (m_lngCurrentToken = UBound(m_Tokens))  ' OR until an error occurs.
    
    obj3DObject = m_obj3DObject
        
    ' Tell user we have finished (which we are, anything after this point is optional)
    ' ================================================================================
    MsgBox "Finished Importing DirectX File.", vbInformation
    
        
    Exit Sub
errTrap:
    Err.Source = mc_ModuleName & ".CreateTokens"
    Call LogAnError(Err)
    
End Sub


Private Sub GetNextToken()
    
    Dim blnValidToken As Boolean
    
    ' Save the existing token.
    m_strOldTokenValue = m_strCurrentTokenValue

    ' =====================================================================
    ' Get the next Token, from the variant array (ie. the Token collection)
    ' =====================================================================
    Do
        blnValidToken = True
        
        ' Increment the token counter and fetch the next token from the array.
        m_lngCurrentToken = m_lngCurrentToken + 1
        m_strCurrentTokenValue = m_Tokens(m_lngCurrentToken)
        
        ' Trim any leading (or trailing) white space (ie. SPACE and TAB characters)
        m_strCurrentTokenValue = TrimWhiteSpace(m_strCurrentTokenValue)
        
        ' Ignore Comment lines starting with // or #, because these comments may contain reserved words.
        If (Left(m_strCurrentTokenValue, 2) = "//" Or Left(m_strCurrentTokenValue, 1) = "#") Then blnValidToken = False
        
        ' Ignore empty tokens
        If Len(m_strCurrentTokenValue) = 0 Then blnValidToken = False
    
    Loop Until (blnValidToken = True) Or (m_lngCurrentToken = UBound(m_Tokens))
    
    
End Sub

Private Function MatchToken(strPattern As String) As Boolean

    On Error GoTo errTrap
    
    ' Assume failure, until proven otherwise.
    MatchToken = False
    
    ' ======================================================================================
    ' Set the following Visual Basic Project Reference:
    '
    '       * Microsoft VBScript Regular Expressions (any version)
    '
    ' "Regular Expressions" are like minature programs that can help us look for
    ' simple or complex patterns - similar to VB's LIKE operator, except much more powerful!
    ' ======================================================================================
    Dim objRegExp As New RegExp
    
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    objRegExp.Pattern = strPattern
    
    ' Test the curent token using a regular expression; ie. strPattern
    MatchToken = objRegExp.Test(m_strCurrentTokenValue)
    
    If MatchToken = True Then Call GetNextToken
    
    Exit Function
errTrap:
    MatchToken = False
        
End Function

Private Function Parse_Template() As Boolean

    On Error GoTo errTrap
    
    Dim strTemplateName As String
    
    ' Assume Failure.
    Parse_Template = False
    
    ' The calling procedure has already verified the format of the
    ' data using the following regular expression:
    '       ^Template\b.{1,}{
    '
    ' We don't have to check it again - just strip it apart to get what we need.
    strTemplateName = TrimWhiteSpace(Mid(m_strOldTokenValue, 9, Len(m_strOldTokenValue) - 9))
    
    'Debug.Print "Template: " & strTemplateName
    
    ' Actually - we're going to completely ignore the "template" template, so just
    ' eat-up the rest of the tokens until we get to the end curly bracket }
    Do
        GetNextToken
    Loop Until MatchToken("^}$") = True
    
    ' Return success.
    Parse_Template = True
    
   Exit Function
errTrap:
    Err.Source = mc_ModuleName & ".Parse_Template"
    Call LogAnError(Err)
    
End Function

Private Function Parse_Header() As Boolean

    On Error GoTo errTrap
    
    ' Assume failure.
    Parse_Header = False
    
    ' The calling procedure has already verified the format of the
    ' data using the following regular expression:
    '       ^Header {$
    '
    
    ' Actually - we're going to completely ignore the "header" template, so just
    ' eat-up the rest of the tokens until we get to the end curly bracket }
    ' Some DirectX files can contain multiple headers. This parser does not deal with that situation.
    Do
        GetNextToken
    Loop Until MatchToken("^}$") = True
    
    'Debug.Print "Header: Found"
    
    ' Return success.
    Parse_Header = True
    
   Exit Function
errTrap:
    Err.Source = mc_ModuleName & ".Parse_Header"
    Call LogAnError(Err)
    
End Function
Private Function Parse_Material() As Boolean

    On Error GoTo errTrap
    
    ' Assume Failure.
    Parse_Material = False
    
    ' The calling procedure has already verified the format of the
    ' data using the following regular expression:
    '       ^Material {$
    '
    
    ' Actually - we're going to completely ignore the "Material" template, so just
    ' eat-up the rest of the tokens until we get to the end curly bracket }
    
    Do
        GetNextToken
    Loop Until MatchToken("^}$") = True
    
    
    'Debug.Print "Material: Found & Ignored."
    
    
    ' Return success.
    Parse_Material = True
    
   Exit Function
errTrap:
    Err.Source = mc_ModuleName & ".Parse_Material"
    Call LogAnError(Err)
    
End Function

Private Function Parse_FrameTransformMatrix() As Boolean

    On Error GoTo errTrap
    
    Dim lngPos As Long
    Dim lngN As Long
    Dim strRows As String
    Dim varMatrix As Variant
    Dim aMatrix As mdrMatrix4
    
    ' Assume failure.
    Parse_FrameTransformMatrix = False
    
    ' This is what we're expecting to see, although it might look slightly different.
    
''    FrameTransformMatrix {
''    0.000000, -0.000000, -1.000000, 0.000000,
''    -1.000000, -0.000000, -0.000000, 0.000000,
''    -0.000000, 1.000000, -0.000000, 0.000000,
''    10.612700, 44.863956, 0.021057, 1.000000;;
''    }


    ' The calling procedure has already verified the format of the
    ' data using the following regular expression:
    '       ^FrameTransformMatrix $
    '
    lngPos = InStr(1, m_strOldTokenValue, "{")
    If lngPos = 0 Then
        ' Advance to the next token if the current token matches the expected curly braket {
        If MatchToken("^{$") <> True Then Err.Raise vbObjectError + 2002, "Parse_FrameTransformMatrix", "'FrameTransformMatrix' template is missing as opening curly bracket {"
    Else
        Call GetNextToken
    End If
    
    'Debug.Print "FrameTransformMatrix: Found"
    
''    ' There should be 4 rows of the matrix - I'm not checking the format, just assuming it's ok.
''    For lngN = 0 To 3
''        strRows = strRows & m_strOldTokenValue
''        Call GetNextToken
''    Next lngN
''    strRows = strRows & ","
    
    strRows = m_strOldTokenValue & ","
    
    ' If you don't know what the Split function does - look it up - very, very, very useful!!!!
    varMatrix = Split(strRows, ",")
    
    ' Load the variant array directly into the Part's Matrix property.
    
    aMatrix.rc11 = CSng(varMatrix(0)): aMatrix.rc12 = CSng(varMatrix(1)): aMatrix.rc13 = CSng(varMatrix(2)): aMatrix.rc14 = CSng(varMatrix(3))
    aMatrix.rc21 = CSng(varMatrix(4)): aMatrix.rc22 = CSng(varMatrix(5)): aMatrix.rc23 = CSng(varMatrix(6)): aMatrix.rc24 = CSng(varMatrix(7))
    aMatrix.rc31 = CSng(varMatrix(8)): aMatrix.rc32 = CSng(varMatrix(9)): aMatrix.rc33 = CSng(varMatrix(10)): aMatrix.rc34 = CSng(varMatrix(11))
    aMatrix.rc41 = CSng(varMatrix(12)): aMatrix.rc42 = CSng(varMatrix(13)): aMatrix.rc43 = CSng(varMatrix(14)): aMatrix.rc44 = CSng(Val(varMatrix(15)))
'    m_obj3DObject.Parts(m_lngPartCount).IdentityMatrix = aMatrix
    m_obj3DObject.Parts(m_lngPartCount).IdentityMatrix = MatrixTranspose(aMatrix)
    'm_obj3DObject.Parts(m_lngPartCount).IdentityMatrix.rc34 = -m_obj3DObject.Parts(m_lngPartCount).IdentityMatrix.rc34 ' Flip z-Axis to convert from Direct X's Left-Handed System, to my Right-Handed System.
    
    ' Return success.
    Parse_FrameTransformMatrix = True
    
   Exit Function
errTrap:
    Err.Source = mc_ModuleName & ".Parse_FrameTransformMatrix"
    Call LogAnError(Err)
    
End Function


Private Function Parse_FileHeader() As Boolean

    On Error GoTo errTrap
        
    '   The File Header information is always on the very first line.
    '   Example:
    '       xof 0302txt 0064
    '
    '   Example2:
    '       xof 0303txt 0032
    '
    '
    Dim strMagicNumber As String    ' 4 Bytes
    Dim strVersionNumber As String  ' 4 Bytes
    Dim strContentType As String    ' 4 Bytes
    Dim strFloatSize As String      ' 4 Bytes
    
    ' Assume Failure.
    Parse_FileHeader = True
    
    ' Read file header information.
    strMagicNumber = Left(m_strOldTokenValue, 4)
    strVersionNumber = Mid(m_strOldTokenValue, 5, 4)
    strContentType = LCase(Mid(m_strOldTokenValue, 9, 4))
    strFloatSize = Mid(m_strOldTokenValue, 13, 4)
    
    'Debug.Print "File Header Version: " & strVersionNumber
    
    If strContentType <> "txt " Then Err.Raise vbObjectError + 2001, "Parse_FileHeader", "Unable to import a DirectX data file of type '" & strContentType & "'"
    
    ' Return success.
    Parse_FileHeader = True
    
   Exit Function
errTrap:
    Err.Source = mc_ModuleName & ".Parse_FileHeader"
    Call LogAnError(Err)
    
End Function


Private Function Parse_Mesh() As Boolean

    ' To make my job easier, this routine assumes a little bit about the format of the data.
    ' There should not be any errors, however if there are, send me an e-mail with the deails
    ' and I'll be happy to help.
    '
    ' support@midar.com
    
    On Error GoTo errTrap
    
    ' Assume Failure.
    Parse_Mesh = False
    
    Dim lngPos As Long
    Dim lngJ As Long
    Dim lngK As Long
    Dim blnMatchFound As Boolean
    
    ' Mesh-specific variables
    Dim strMeshName As String
    Dim lngNumVertices As Long
    Dim lngNumFaces As Long
    Dim lngNumFaceVertexIndices As Long
    Dim varVertices As Variant
    Dim varFaceData As Variant
    Dim varVertexIndices As Variant
    
    ' The calling procedure has already verified the **partial** format
    ' of the "Frame" template using the following regular expression:
    '       ^Mesh {$
    '
    lngPos = InStr(1, m_strOldTokenValue, "{")
    If lngPos > 0 Then
        strMeshName = TrimWhiteSpace(Mid(m_strOldTokenValue, 6, lngPos - 6))
    Else
        ' The beginning curly bracket { is not on this line, most probably on the next.
        strMeshName = TrimWhiteSpace(Mid(m_strOldTokenValue, 6))
        
        ' Advance to the next token if the current token matches the expected curly braket {
        If MatchToken("^{$") <> True Then Err.Raise vbObjectError + 2002, "Parse_Mesh", "'Mesh' template '" & strMeshName & "' is missing as opening curly bracket {"
    End If
    
    If strMeshName = "" Then strMeshName = "unnamed_mesh"
    
    'Debug.Print "Mesh: " & strMeshName
    
    ' Get Number of Mesh Vertices
    ' ===========================
    If MatchToken("^\d{1,};$") = True Then
        lngNumVertices = Val(m_strOldTokenValue) ' Val automatically strips off the trailing semicolon
        
        'Debug.Print "Vertices: " & lngNumVertices
                    
        ReDim m_obj3DObject.Parts(m_lngPartCount).Vertices(lngNumVertices - 1)
        m_obj3DObject.Parts(m_lngPartCount).Caption = strMeshName
        

        ' Load the Mesh Vertices in X, Y, Z format
        ' ========================================
        For lngJ = 0 To lngNumVertices - 1
            Call GetNextToken ' ie. Each XYZ row of vertices SHOULD be on a separate line, I'm not checking!
            varVertices = Split(m_strOldTokenValue, ";")
            
            'Debug.Print "X: " & varVertices(0) & "   Y: " & varVertices(1) & "   Z: " & varVertices(2)
            
            With m_obj3DObject.Parts(m_lngPartCount).Vertices(lngJ).Pxyz
                .x = varVertices(0)
                .y = varVertices(1)
                .Z = varVertices(2)
                .w = 1
            End With
            
        Next lngJ
        
        ' Get Number of Mesh Faces
        ' ========================
        If MatchToken("^\d{1,};$") = True Then
            lngNumFaces = Val(m_strOldTokenValue) ' Val automatically strips off the trailing semicolon
            
            'Debug.Print "Faces: " & lngNumFaces
            
            ReDim m_obj3DObject.Parts(m_lngPartCount).Faces(lngNumFaces - 1)
            
            ' Process the Mesh Faces
            ' ======================
            For lngJ = 0 To lngNumFaces - 1
                  
                'Debug.Print "Face " & lngJ;
                
                Call GetNextToken ' ie. Each "Face" should be on a separate row, but I'm not checking!
                
                varFaceData = Split(m_strOldTokenValue, ";")
                
                lngNumFaceVertexIndices = Val(varFaceData(0)) ' Val automatically strips off the trailing semicolon
                
                'Debug.Print " with " & lngNumFaceVertexIndices & " Face Vertex Indices: ";
                
                ' Load the Face Vertex Indices
                ' ============================
                varVertexIndices = Split(varFaceData(1), ",")
                
                'Debug.Print varFaceData(1)
                
                m_obj3DObject.Parts(m_lngPartCount).Faces(lngJ) = varVertexIndices

            Next lngJ
            'Debug.Print
        Else ' (Get Number of Mesh Faces)
            Err.Raise vbObjectError + 2001, "Parse_Mesh", "'Mesh' template has a missing or invalid number of Faces, '" & m_strCurrentTokenValue & "'"
        End If
    Else ' (Get Number of Mesh Vertices)
        Err.Raise vbObjectError + 2001, "Parse_Mesh", "'Mesh' template has a missing or invalid number of Vertices, '" & m_strCurrentTokenValue & "'"
    End If
    
    
    ' Return success.
    Parse_Mesh = True
    
   Exit Function
errTrap:
    Err.Source = mc_ModuleName & ".Parse_Mesh"
    Call LogAnError(Err)
    
End Function

Private Function Parse_Frame() As Boolean

    On Error GoTo errTrap
    
    Dim strFrameName As String
    Dim lngPos As Long
    Dim blnMatchFound As Boolean
    Dim blnMeshDone As Boolean
    
    ' Assume Failure.
    Parse_Frame = False
    
    ' The calling procedure has already verified the **partial** format
    ' of the "Frame" template using the following regular expression:
    '       ^Frame .{1,}
    ' Now look for the opening curly bracket.
    
    lngPos = InStr(1, m_strOldTokenValue, "{")
    If lngPos > 0 Then
        strFrameName = TrimWhiteSpace(Mid(m_strOldTokenValue, 7, lngPos - 7))
    Else
        ' The beginning curly bracket { is not on this line, most probably on the next.
        
        strFrameName = TrimWhiteSpace(Mid(m_strOldTokenValue, 7))
        
        ' Advance to the next token if the current token matches the expected curly braket {
        If MatchToken("^{$") <> True Then Err.Raise vbObjectError + 2002, "Parse_Frame", "'Frame' template '" & strFrameName & "' is missing as opening curly bracket {"
        
    End If
    
    If strFrameName = "" Then strFrameName = "unnamed_frame"

    'Debug.Print "Frame: " & strFrameName
    
    m_lngPartCount = m_lngPartCount + 1 ' Note: Set to minus one (-1) when module starts.
    ReDim Preserve m_obj3DObject.Parts(m_lngPartCount)
    m_obj3DObject.Parts(m_lngPartCount).Caption = strFrameName
    
    
    ' A FRAME template can only contain the following objects:
    '   * Frame (ie. recursivly)
    '   * FrameTransformMatrix
    '   * Mesh
    
    Do
        blnMeshDone = False
        blnMatchFound = False

        If MatchToken("^Frame .{1,}") = True Then
        
            blnMatchFound = Parse_Frame() ' << Putting the empty brackets here is very important because within this context, "Parse_Frame" is the name of a variable, and also a Function.
            
        ElseIf MatchToken("^FrameTransformMatrix .{1,}") = True Then

            blnMatchFound = Parse_FrameTransformMatrix()
            
        ElseIf MatchToken("^Mesh .{1,}") = True Then
            
            blnMeshDone = Parse_Mesh()
            
        End If
        
        ' This little bit helps us skip ahead to the next
        ' Frame template (if any) after we've got the Mesh.
        If (blnMeshDone = False) And (blnMatchFound = False) Then
            blnMatchFound = True
            Call GetNextToken
        End If
        
    Loop Until (blnMeshDone = True) Or (blnMatchFound = False) Or (m_lngCurrentToken = UBound(m_Tokens))  ' OR until an error occurs.
    
    
    ' Return success.
    Parse_Frame = True
    
    Exit Function
errTrap:
    Err.Source = mc_ModuleName & ".Parse_Frame"
    Call LogAnError(Err)
    
End Function


