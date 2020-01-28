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

' Finds the next instance of a string in the text field. (Pass empty string to continue searching for the last string)
Public Sub FindNext(sWhat As String)
    Dim sNeedle As String
    Dim sHaystack As String
    
    ' If the query is different than Null, use it to begin the search.
    If sWhat <> "" Then
        sSearchString = sWhat
    End If
    
    ' Check for invalid search strings.
    If sSearchString = "" Then
        MsgBox "Search query is empty", vbOKOnly + vbExclamation, "Empty Search"
        Exit Sub
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
    Else
        ' Nothing was found.
        MsgBox "Cannot find '" & sSearchString & "'", vbOKOnly + vbInformation, "Search Failed"
        bFindFoundSomething = False
    End If
End Sub
