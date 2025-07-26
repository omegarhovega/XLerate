' clsDynamicButtonHandler
Option Explicit

Public WithEvents DynamicButton As MSForms.CommandButton
Private ButtonRole As String
Private ParentForm As Object

Public Sub Initialize(ByRef Button As MSForms.CommandButton, ByVal Role As String, ByRef Parent As Object)
    Debug.Print "Initializing button handler for role: " & Role
    Set DynamicButton = Button
    ButtonRole = Role
    Set ParentForm = Parent
    Debug.Print "Button handler initialized"
End Sub

Private Sub DynamicButton_Click()
    Debug.Print "DynamicButton_Click triggered for role: " & ButtonRole
    On Error GoTo ErrorHandler
    
    Select Case ButtonRole
        Case "Add"
            Debug.Print "Calling btnAdd_Click"
            ParentForm.btnAdd_Click
        Case "Remove"
            Debug.Print "Calling btnRemove_Click"
            ParentForm.btnRemove_Click
        Case "Save"
            Debug.Print "Calling btnSave_Click"
            ParentForm.btnSave_Click
        Case "Reset"
            Debug.Print "Calling btnReset_Click"
            ParentForm.btnReset_Click
        Case "Cancel"
            Debug.Print "Calling btnCancel_Click"
            ParentForm.btnCancel_Click
        Case "BorderColor"
            Debug.Print "Calling btnBorderColor_Click"
            ParentForm.btnBorderColor_Click
        Case "FillColor"
            Debug.Print "Calling btnFillColor_Click"
            ParentForm.btnFillColor_Click
        Case "FontColor"
            Debug.Print "Calling btnFontColor_Click"
            ParentForm.btnFontColor_Click
    End Select
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in DynamicButton_Click: " & Err.Description & " (Error " & Err.Number & ")"
    Resume Next
End Sub
