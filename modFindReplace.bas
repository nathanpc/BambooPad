Attribute VB_Name = "modFindReplace"
Option Explicit

' Direction contants.
Public Const frUp = 0
Public Const frDown = 1
Public Const frAny = 2

' General variables.
Private txtSearch As TextBox
Public bFindFoundSomething As Boolean
Public bFindMatchCase As Boolean
Public fFindDirection As Integer
Private nLastPosition As Long
Private sSearchString As String

' Initializes the find and replace module.
Public Sub InitializeFindReplace(ByRef txField As TextBox, bMatchCase As Boolean, fDirection As Integer)
    Set txtSearch = txField
    bFindMatchCase = bMatchCase
    fFindDirection = fDirection
    nLastPosition = 0
    sSearchString = ""
    bFindFoundSomething = False
End Sub

' Finds and selects a string in a text. Returns True if something was found.
Private Function FindAndSelectText(sWhat As String) As Boolean
    Dim sNeedle As String
    Dim sHaystack As String
    
    ' If the query is different than Null, use it to begin the search.
    If sWhat <> "" Then
        sSearchString = sWhat
    End If
    
    ' Check for invalid search strings.
    If sSearchString = "" Then
        MsgBox "Search query is empty", vbOKOnly + vbExclamation, "Empty Search"
        FindAndSelectText = False
        Exit Function
    End If
    
    ' Set query and searched strings and matches the cases if needed.
    sNeedle = sSearchString
    sHaystack = txtSearch.Text
    If Not bFindMatchCase Then
        sNeedle = LCase(sNeedle)
        sHaystack = LCase(sHaystack)
    End If
    
    ' Set cursor start position.
    If fFindDirection = frAny Then
        nLastPosition = 0
    ElseIf fFindDirection = frDown Then
        nLastPosition = txtSearch.SelStart + txtSearch.SelLength
    ElseIf fFindDirection = frUp Then
        nLastPosition = txtSearch.SelStart
    End If
    
    ' Just make sure for InStr.
    If nLastPosition = 0 Then
        nLastPosition = 1
    End If
    
    ' Perform search.
    If fFindDirection = frUp Then
        nLastPosition = InStrRev(sHaystack, sNeedle, nLastPosition, vbBinaryCompare)
    Else
        nLastPosition = InStr(nLastPosition, sHaystack, sNeedle, vbBinaryCompare)
    End If
    
    If nLastPosition <> 0 Then
        ' Found something.
        txtSearch.SelStart = nLastPosition - 1
        txtSearch.SelLength = Len(sNeedle)
        bFindFoundSomething = True
        
        ' Continue searching if using Any direction.
        If fFindDirection = frAny Then
            fFindDirection = frDown
        End If
        
        FindAndSelectText = True
    Else
        ' Nothing was found.
        bFindFoundSomething = False
        FindAndSelectText = False
    End If
End Function

' Finds the next instance of a string in the text field. (Pass empty string to continue searching for the last string)
Public Sub FindNext(sWhat As String)
    If FindAndSelectText(sWhat) Then
        txtSearch.SetFocus
    Else
        MsgBox "Cannot find '" & sSearchString & "'", vbOKOnly + vbInformation, "Search Failed"
    End If
End Sub

' Checks if the text field is ready for a replace operation.
Private Function IsPreparedForReplace() As Boolean
    IsPreparedForReplace = (txtSearch.SelLength > 0)
End Function

' Replaces the selected text with the provided string.
Private Sub ReplaceSelected(sWith As String)
    ' Abort if there's no text selected.
    If Not IsPreparedForReplace Then
        Exit Sub
    End If
    
    txtSearch.SelText = sWith
End Sub

' Finds and replace a single instance of a text with a string. Returns True if it did find something.
Private Function ReplaceTextOnce(sWhat As String, sWith As String) As Boolean
    If FindAndSelectText(sWhat) Then
        ' Found something.
        ReplaceSelected sWith
        ReplaceTextOnce = True
    Else
        ' No luck.
        ReplaceTextOnce = False
    End If
End Function

' Replaces a single instace of a string within a text.
Public Sub ReplaceOnce(sWhat As String, sWith As String)
    If Not ReplaceTextOnce(sWhat, sWith) Then
        MsgBox "Cannot find '" & sSearchString & "'", vbOKOnly + vbInformation, "Search Failed"
    End If
End Sub

' Replaces all instaces of a string within a text.
Public Sub ReplaceAll(sWhat As String, sWith As String)
    Do While ReplaceTextOnce(sWhat, sWith)
    Loop
End Sub
