Attribute VB_Name = "modUndoRedo"
' UndoRedo
' Undo/Redo stack helper module.
'
' Author: Nathan Campos <nathan@innoveworkshop.com>

Option Explicit

' State variables.
Const nMaxUndo = 10  ' Remember to change asUndoStack and anCursorStack value too.
Private asUndoStack(10) As String
Private anCursorStack(10) As Integer
Private nUndoIndex As Integer
Private nLastUndoIndex As Integer

' Control objects.
Private btnUndo As CommandBarButton
Private btnRedo As CommandBarButton
Private mniUndo As CommandbarLib.Item
Private mniRedo As CommandbarLib.Item
Private txtField As TextBox

' Initializes the stack and gets references to the controls we need to operate with.
Public Sub InitializeUndoStack(ByRef btUndo As CommandBarButton, ByRef btRedo As CommandBarButton, ByRef miUndo As CommandbarLib.Item, ByRef miRedo As CommandbarLib.Item, ByRef txField As TextBox, ByVal sInitialText As String)
    ' Initialize the counters.
    nUndoIndex = 0
    nLastUndoIndex = 0
    asUndoStack(0) = sInitialText
    anCursorStack(0) = 1
    
    ' Set the control references.
    Set btnUndo = btUndo
    Set btnRedo = btRedo
    Set mniUndo = miUndo
    Set mniRedo = miRedo
    Set txtField = txField
    
    ' Disable both to start.
    SetUndoState False
    SetRedoState False
End Sub

' Sets the current undo stack item, before a shift operation happens.
Public Sub SetCurrentUndoStackItem(sText As String, nCursorPosition As Integer)
    asUndoStack(nUndoIndex) = sText
    anCursorStack(nUndoIndex) = nCursorPosition
    nLastUndoIndex = nUndoIndex
    
    SetControlsState
End Sub

' Pushes text into the undo stack.
Public Sub PushTextToUndoStack(sText As String, nCursorPosition As Integer)
    ' Check if we should shift the stack before pushing the new item.
    Dim nShiftCounter As Integer
    If nUndoIndex = nMaxUndo Then
        For nShiftCounter = 1 To 10
            asUndoStack(nShiftCounter - 1) = asUndoStack(nShiftCounter)
            anCursorStack(nShiftCounter - 1) = anCursorStack(nShiftCounter)
        Next nShiftCounter
    End If
    
    ' Set stack state.
    asUndoStack(nUndoIndex) = sText
    anCursorStack(nUndoIndex) = nCursorPosition
    
    ' Operate the index.
    If nUndoIndex < nMaxUndo Then
        nUndoIndex = nUndoIndex + 1
        nLastUndoIndex = nUndoIndex
    End If
    
    SetControlsState
End Sub

' Gets text from the undo stack and moves the index backwards.
Public Function GetTextFromUndoStack() As String
    If nUndoIndex > 0 Then
        nUndoIndex = nUndoIndex - 1
    End If
    
    GetTextFromUndoStack = asUndoStack(nUndoIndex)
    txtField.SelStart = anCursorStack(nUndoIndex)
    SetControlsState
End Function

' Gets text from the undo stack backwards. Used for redo.
Public Function GetTextFromRedoStack() As String
    If (nUndoIndex < nMaxUndo) And (nUndoIndex < nLastUndoIndex) Then
        nUndoIndex = nUndoIndex + 1
    End If
    
    GetTextFromRedoStack = asUndoStack(nUndoIndex)
    txtField.SelStart = anCursorStack(nUndoIndex)
    SetControlsState
End Function

' Clears the stack.
Public Sub ClearUndoStack(strInitialValue As String)
    ' Reset the counters.
    nUndoIndex = 0
    nLastUndoIndex = 0
    
    ' Reset the first value in the stack.
    SetCurrentUndoStackItem strInitialValue, 1
End Sub

' Sets the controls Enabled state according to the index counts.
Private Sub SetControlsState()
    If (nUndoIndex < nLastUndoIndex) And (nUndoIndex > 0) Then
        SetUndoState True
        SetRedoState True
    ElseIf (nUndoIndex = nLastUndoIndex) And (nUndoIndex > 0) Then
        SetUndoState True
        SetRedoState False
    ElseIf (nUndoIndex < nLastUndoIndex) And (nUndoIndex = 0) Then
        SetUndoState False
        SetRedoState True
    Else
        SetUndoState False
        SetUndoState False
    End If
    
    ' Bring the focus back to the text field.
    txtField.SetFocus
End Sub

' Sets the Enabled state of the undo controls.
Private Sub SetUndoState(bEnabled As Boolean)
    If btnUndo.Enabled <> bEnabled Then
        btnUndo.Enabled = bEnabled
        mniUndo.Enabled = bEnabled
    End If
End Sub

' Sets the Enabled state of the redo controls.
Private Sub SetRedoState(bEnabled As Boolean)
    If btnRedo.Enabled <> bEnabled Then
        btnRedo.Enabled = bEnabled
        mniRedo.Enabled = bEnabled
    End If
End Sub
