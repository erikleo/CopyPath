Attribute VB_Name = "CopyPath"
Sub Main()
    If Command() = "" Then
        MsgBox "no argument"
        End
    End If
    
    Dim myPath As String
    Clipboard.Clear
    myPath = CleanPathString(Command())
    
    Dim e As Variant
    Dim curPath As Variant
    Dim uncServer As String
    
    e = Split(myPath, vbNewLine)
    myPath = ""
    For Each curPath In e
        If curPath <> "" Then
            ''''''''''''''for windows nt only
            If InStr(curPath, "~") > 0 And InStr(curPath, "\\") = 0 Then
                curPath = GetLongFilename(curPath)
            End If
            '''''''''''''''
            'get server instead of drive letter
            If (Mid(curPath, 2, 2) = ":\") Then
                uncServer = LetterToUNC(Left(curPath, 2))
                curPath = Right(curPath, Len(curPath) - 2)
                myPath = myPath & uncServer & curPath & vbNewLine
            Else
                myPath = myPath & curPath & vbNewLine
            End If
        End If
    Next
    'remove last newline
    myPath = Left(myPath, Len(myPath) - 1)
    Clipboard.SetText myPath
End Sub
Function checkQuotes(myString As String) As String
    Dim Result
    Dim shorter
    If Asc(Left(myString, 1)) = 34 Then
        shorter = Left(myString, Len(myString) - 1)
        Result = Right(shorter, Len(myString) - 2)
        'put muliple files on new lines
        checkQuotes = Replace(Result, """ """, vbNewLine)
    Else
    'if there are not quotes in the string then spaces mean a new file so put on a new line
    checkQuotes = Replace(myString, " ", vbNewLine)
    End If
    
End Function

Function CleanPathString(ByRef inPath As String) As String
    
    Dim startChar As String
    Dim endChar As String
    Dim newStartChar As String
    
    If inPath = "" Then
        CleanPathString = ""
    Else
        If Asc(Left(inPath, 1)) = 34 Then
            startChar = 2
            endChar = InStr(startChar, inPath, """") - 2
            newStartChar = endChar + 4
        Else
            startChar = 1
            endChar = InStr(inPath, " ") - 1
            newStartChar = endChar + 2
        End If
        If endChar = -1 Then
            CleanPathString = inPath
            Exit Function
        End If
    CleanPathString = Mid(inPath, startChar, endChar) & vbNewLine & CleanPathString(Trim(Mid(inPath, newStartChar)))
    End If
End Function


'used for Windows NT
Function GetLongFilename(ByVal sShortName As String) As String
    Dim sLongName As String
    Dim sTemp As String
    Dim iSlashPos As Integer
    
    'Add \ to short name to prevent Instr from failing
    sShortName = sShortName & "\"
    
    'Start from 4 to ignore the "[Drive Letter]:\" characters
    iSlashPos = InStr(4, sShortName, "\")
    
    'Pull out each string between \ character for conversion
    While iSlashPos
        sTemp = Dir(Left$(sShortName, iSlashPos - 1), _
        vbNormal + vbHidden + vbSystem + vbDirectory)
        If sTemp = "" Then
            'Error 52 - Bad File Name or Number
            GetLongFilename = ""
            Exit Function
        End If
        sLongName = sLongName & "\" & sTemp
        iSlashPos = InStr(iSlashPos + 1, sShortName, "\")
    Wend
    
    'Prefix with the drive letter
    GetLongFilename = Left$(sShortName, 2) & sLongName
End Function
