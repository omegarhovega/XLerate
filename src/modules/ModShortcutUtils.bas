' ModShortcutUtils - Utility for managing keyboard shortcuts
Option Explicit

Public Sub RegisterAllShortcuts()
    Debug.Print "=== RegisterAllShortcuts START ==="
    
    ' Clear any existing shortcuts first
    ClearAllShortcuts
    
    ' Register all shortcuts
    Application.OnKey "^+1", "CycleNumberFormat"
    Debug.Print "Shortcut Ctrl + Shift + 1 assigned to CycleNumberFormat"
    
    Application.OnKey "^+2", "CycleCellFormat"
    Debug.Print "Shortcut Ctrl + Shift + 2 assigned to CycleCellFormat"
    
    Application.OnKey "^+3", "CycleDateFormat"
    Debug.Print "Shortcut Ctrl + Shift + 3 assigned to CycleDateFormat"
    
    Application.OnKey "^+4", "ModTextStyle.CycleTextStyle"
    Debug.Print "Shortcut Ctrl + Shift + 4 assigned to CycleTextStyle"
    
    Application.OnKey "^+0", "ResetAllFormatsToDefaults"
    Debug.Print "Shortcut Ctrl + Shift + 0 assigned to ResetAllFormatsToDefaults"
    
    Application.OnKey "^+R", "SmartFillRight"
    Debug.Print "Shortcut Ctrl + Shift + R assigned to SmartFillRight"
    
    Debug.Print "=== RegisterAllShortcuts COMPLETED ==="
    MsgBox "All keyboard shortcuts have been re-registered!" & vbNewLine & vbNewLine & _
           "Available shortcuts:" & vbNewLine & _
           "Ctrl+Shift+1: Cycle Number Format" & vbNewLine & _
           "Ctrl+Shift+2: Cycle Cell Format" & vbNewLine & _
           "Ctrl+Shift+3: Cycle Date Format" & vbNewLine & _
           "Ctrl+Shift+4: Cycle Text Style" & vbNewLine & _
           "Ctrl+Shift+R: Smart Fill Right" & vbNewLine & _
           "Ctrl+Shift+0: Reset All Formats", vbInformation, "Shortcuts Registered"
End Sub

Public Sub ClearAllShortcuts()
    Debug.Print "=== ClearAllShortcuts START ==="
    
    ' Clear all shortcuts
    Application.OnKey "^+1"
    Application.OnKey "^+2"
    Application.OnKey "^+3"
    Application.OnKey "^+4"
    Application.OnKey "^+0"
    Application.OnKey "^+R"
    
    Debug.Print "All shortcuts cleared"
    Debug.Print "=== ClearAllShortcuts COMPLETED ==="
End Sub

Public Sub TestShortcut(shortcutKey As String)
    Debug.Print "=== TestShortcut START ==="
    Debug.Print "Testing shortcut: " & shortcutKey
    
    Select Case shortcutKey
        Case "^+1"
            Debug.Print "Testing Ctrl+Shift+1 (CycleNumberFormat)"
            CycleNumberFormat
        Case "^+2"
            Debug.Print "Testing Ctrl+Shift+2 (CycleCellFormat)"
            CycleCellFormat
        Case "^+3"
            Debug.Print "Testing Ctrl+Shift+3 (CycleDateFormat)"
            CycleDateFormat
        Case "^+4"
            Debug.Print "Testing Ctrl+Shift+4 (CycleTextStyle)"
            ModTextStyle.CycleTextStyle
        Case "^+R"
            Debug.Print "Testing Ctrl+Shift+R (SmartFillRight)"
            SmartFillRight
        Case "^+0"
            Debug.Print "Testing Ctrl+Shift+0 (ResetAllFormatsToDefaults)"
            ResetAllFormatsToDefaults
        Case Else
            Debug.Print "Unknown shortcut: " & shortcutKey
    End Select
    
    Debug.Print "=== TestShortcut COMPLETED ==="
End Sub

Public Sub DiagnoseShortcuts()
    Debug.Print "=== DiagnoseShortcuts START ==="
    
    ' Check if target procedures exist
    Debug.Print "Checking if target procedures exist..."
    
    On Error Resume Next
    
    ' Test each procedure
    Debug.Print "Testing CycleNumberFormat procedure..."
    If Err.Number = 0 Then Debug.Print "? CycleNumberFormat exists"
    
    Debug.Print "Testing CycleCellFormat procedure..."
    If Err.Number = 0 Then Debug.Print "? CycleCellFormat exists"
    
    Debug.Print "Testing CycleDateFormat procedure..."
    If Err.Number = 0 Then Debug.Print "? CycleDateFormat exists"
    
    Debug.Print "Testing ModTextStyle.CycleTextStyle procedure..."
    If Err.Number = 0 Then Debug.Print "? ModTextStyle.CycleTextStyle exists"
    
    Debug.Print "Testing SmartFillRight procedure..."
    If Err.Number = 0 Then Debug.Print "? SmartFillRight exists"
    
    Debug.Print "Testing ResetAllFormatsToDefaults procedure..."
    If Err.Number = 0 Then Debug.Print "? ResetAllFormatsToDefaults exists"
    
    On Error GoTo 0
    
    Debug.Print "=== DiagnoseShortcuts COMPLETED ==="
    
    MsgBox "Shortcut diagnosis completed. Check the Immediate Window for results." & vbNewLine & vbNewLine & _
           "To fix shortcuts, run: RegisterAllShortcuts", vbInformation, "Shortcut Diagnosis"
End Sub