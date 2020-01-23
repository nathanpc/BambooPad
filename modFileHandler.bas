Attribute VB_Name = "modFileHandler"
' FileHandler
' File handling helper module.
'
' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' Slurps the contents of a file into a string.
Public Function GetFileText(hndFile As File, sFilename As String, bConvertLFtoCRLF As Boolean) As String
    Dim sBuffer As String
    
    If bConvertLFtoCRLF Then
        MsgBox "TODO: Implement LF to CRLF conversion"
    End If
      
    ' Slurp file into buffer.
    hndFile.Open sFilename, fsModeInput, fsAccessRead, fsLockRead
    If Not hndFile.EOF Then
        sBuffer = hndFile.Input(hndFile.LOF)
    End If
    hndFile.Close
    
    GetFileText = sBuffer
End Function

' Saves the contents of a string to a file.
Public Sub WriteFileText(hndFile As File, sFilename As String, sContents As String, bConvertCRLFtoLF As Boolean)
    If bConvertCRLFtoLF Then
        MsgBox "TODO: Implement CRLF to LF conversion"
    End If
    
    hndFile.Open sFilename, fsModeOutput, fsAccessReadWrite, fsLockReadWrite
    hndFile.LinePrint sContents
    hndFile.Close
End Sub

' Check if a file is using DOS or Unix line endings.
Public Function IsFileUsingCRLF(sText As String) As Boolean
    ' TODO: Search for a single occurence of \r
    IsFileUsingCRLF = True
End Function

' Converts from DOS line endings to Unix.
Public Function ConvertCRLFtoLF(sText As String) As String
    ' TODO: Go through the text and remove any \r.
    ConvertCRLFtoLF = sText
End Function

' Converts from Unix line endings to DOS.
Public Function ConvertLFtoCRLF(sText As String) As String
    ' TODO: Go through the text and replace every \n with \r\n
    ConvertLFtoCRLF = sText
End Function

' Gets the character position of the beginning of a line numnber. Returns negative if not found.
Public Function GetLineCharPosition(sText As String, nLine As Integer) As Integer
    Dim nLineCount As Integer
    Dim nCharCount As Integer
    nLineCount = 1
    
    ' If all you want is the first line, then do nothing.
    If nLine > 1 Then
        nCharCount = InStr(1, sText, vbLf)
        
        Do While nCharCount > 0
            nLineCount = nLineCount + 1
            If nLineCount = nLine Then
                GetLineCharPosition = nCharCount
                Exit Function
            End If
            
            nCharCount = InStr(nCharCount + 1, sText, vbLf)
        Loop
    Else
        GetLineCharPosition = 0
        Exit Function
    End If
    
    GetLineCharPosition = -1
End Function
