Option Explicit

Private Function fullAddress(inCell As Range) As String
    fullAddress = inCell.address(External:=True)
End Function

Public Sub ShowTracePrecedents()
    ' Ensure user has selected a range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell or range."
        Exit Sub
    End If
    
    Dim selectedRange As Range
    Set selectedRange = Selection.Cells(1, 1) ' Use only the first cell for recursive tracing
    
    ' Create the UserForm
    Dim frmPrecedents As New frmPrecedents
    With frmPrecedents
        ' Clear existing items
        .lstPrecedents.Clear
        
        ' Set up list box properties
        .lstPrecedents.ColumnCount = 3
        
        Dim columnWidths As String
        columnWidths = .GetColumnWidths()
        .lstPrecedents.ColumnWidths = columnWidths
        
        ' Add headers using separate labels
        .AddHeaders
        
        ' Display formula of selected cell
        .Controls("txtFormula").text = selectedRange.formula
        
        ' Add the root cell first
        .lstPrecedents.AddItem
        
        With .lstPrecedents
            Dim rootAddress As String
            rootAddress = selectedRange.Worksheet.Name & "!" & selectedRange.address
            
            .List(.ListCount - 1, 0) = rootAddress
            
            Dim rootValue As String
            rootValue = GetCellValueAsString(selectedRange)
            .List(.ListCount - 1, 1) = rootValue
            
            .List(.ListCount - 1, 2) = selectedRange.formula
        End With
    
        ' Use new recursive logic to get hierarchical precedents
        Dim processedCells As New Collection
        
        Dim precedentsStr As String
        Dim maxDepth As Integer
        maxDepth = Val(GetSetting("XLerate", "FormulaTracing", "MaxDepth", "10"))
        
        precedentsStr = GetReferencesRecursive(selectedRange, True, 1, maxDepth, processedCells)
        
        ' Parse and populate the hierarchical precedents
        If Len(precedentsStr) > 0 Then
            Dim precedentLines() As String
            precedentLines = Split(precedentsStr, vbCr)
            
            Dim i As Long
            For i = 0 To UBound(precedentLines)
                
                If Len(Trim(precedentLines(i))) > 0 Then
                    Dim line As String
                    line = precedentLines(i)
                    
                    ' Extract level and address from formatted line (e.g., "  L1: [Sheet1]$A$1")
                    Dim colonPos As Long
                    colonPos = InStr(line, ": ")
                    
                    If colonPos > 0 Then
                        Dim indentAndLevel As String
                        Dim precedentAddress As String
                        indentAndLevel = Left(line, colonPos - 1)
                        precedentAddress = Mid(line, colonPos + 2)
                        
                        ' Clean up the address - remove any extra text like "(from range...)"
                        Dim cleanAddress As String
                        Dim parenPos As Long
                        parenPos = InStr(precedentAddress, " (from range")
                        
                        If parenPos > 0 Then
                            cleanAddress = Trim(Left(precedentAddress, parenPos - 1))
                        Else
                            cleanAddress = Trim(precedentAddress)
                        End If
                        
                        ' Get the actual range object for the precedent
                        Dim precedentRange As Range
                        On Error Resume Next
                        Set precedentRange = Range(cleanAddress)
                        Dim rangeError As Long
                        rangeError = Err.Number
                        
                        If Not precedentRange Is Nothing And rangeError = 0 Then
                        Else
                        End If
                        
                        .lstPrecedents.AddItem
                        
                        With .lstPrecedents
                            .List(.ListCount - 1, 0) = indentAndLevel & ": " & precedentAddress
                            
                            If Not precedentRange Is Nothing And rangeError = 0 Then
                                Dim cellValue As String
                                cellValue = GetCellValueAsString(precedentRange)
                                .List(.ListCount - 1, 1) = cellValue
                                
                                .List(.ListCount - 1, 2) = precedentRange.formula
                            Else
                                .List(.ListCount - 1, 1) = "#N/A"
                                .List(.ListCount - 1, 2) = "#N/A"
                            End If
                        End With
                        On Error GoTo 0
                    End If
                Else
                End If
            Next i
        End If

        ' Select the first row if available
        If .lstPrecedents.ListCount > 0 Then
            .lstPrecedents.ListIndex = 0
        Else
        End If
        
        ' Convert to tree view with expand/collapse functionality
        .ConvertToTreeView
        .Show vbModeless
    End With
End Sub

Public Sub ShowTraceDependents()
    ' Ensure user has selected a range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell or range."
        Exit Sub
    End If
    
    Dim selectedRange As Range
    Set selectedRange = Selection.Cells(1, 1) ' Use only the first cell for recursive tracing
    
    ' Create the UserForm
    Dim frmDependents As New frmDependents
    With frmDependents
        ' Clear existing items
        .lstDependents.Clear
        
        ' Set up list box properties
        .lstDependents.ColumnCount = 3
        
        Dim columnWidths As String
        columnWidths = .GetColumnWidths()
        .lstDependents.ColumnWidths = columnWidths
        
        ' Add headers using separate labels
        .AddHeaders
        
        ' Display formula of selected cell
        .Controls("txtFormula").text = selectedRange.formula
        
        ' Add the root cell first
        .lstDependents.AddItem
        
        With .lstDependents
            Dim rootAddress As String
            rootAddress = selectedRange.Worksheet.Name & "!" & selectedRange.address
            
            .List(.ListCount - 1, 0) = rootAddress
            
            Dim rootValue As String
            rootValue = GetCellValueAsString(selectedRange)
            .List(.ListCount - 1, 1) = rootValue
            
            .List(.ListCount - 1, 2) = selectedRange.formula
        End With
    
        ' Use new recursive logic to get hierarchical dependents
        Dim processedCells As New Collection
        
        Dim dependentsStr As String
        Dim maxDepth As Integer
        maxDepth = Val(GetSetting("XLerate", "FormulaTracing", "MaxDepth", "10"))
        
        dependentsStr = GetReferencesRecursive(selectedRange, False, 1, maxDepth, processedCells)
        
        ' Parse and populate the hierarchical dependents
        If Len(dependentsStr) > 0 Then
            Dim dependentLines() As String
            dependentLines = Split(dependentsStr, vbCr)
            
            Dim i As Long
            For i = 0 To UBound(dependentLines)
                
                If Len(Trim(dependentLines(i))) > 0 Then
                    Dim line As String
                    line = dependentLines(i)
                    
                    ' Extract level and address from formatted line (e.g., "  L1: [Sheet1]$A$1")
                    Dim colonPos As Long
                    colonPos = InStr(line, ": ")
                    
                    If colonPos > 0 Then
                        Dim indentAndLevel As String
                        Dim dependentAddress As String
                        indentAndLevel = Left(line, colonPos - 1)
                        dependentAddress = Mid(line, colonPos + 2)
                        
                        ' Clean up the address - remove any extra text like "(from range...)"
                        Dim cleanAddress As String
                        Dim parenPos As Long
                        parenPos = InStr(dependentAddress, " (from range")
                        
                        If parenPos > 0 Then
                            cleanAddress = Trim(Left(dependentAddress, parenPos - 1))
                        Else
                            cleanAddress = Trim(dependentAddress)
                        End If
                        
                        ' Get the actual range object for the dependent
                        Dim dependentRange As Range
                        On Error Resume Next
                        Set dependentRange = Range(cleanAddress)
                        Dim rangeError As Long
                        rangeError = Err.Number
                        
                        If Not dependentRange Is Nothing And rangeError = 0 Then
                        Else
                        End If
                        
                        .lstDependents.AddItem
                        
                        With .lstDependents
                            .List(.ListCount - 1, 0) = indentAndLevel & ": " & dependentAddress
                            
                            If Not dependentRange Is Nothing And rangeError = 0 Then
                                Dim cellValue As String
                                cellValue = GetCellValueAsString(dependentRange)
                                .List(.ListCount - 1, 1) = cellValue
                                
                                .List(.ListCount - 1, 2) = dependentRange.formula
                            Else
                                .List(.ListCount - 1, 1) = "#N/A"
                                .List(.ListCount - 1, 2) = "#N/A"
                            End If
                        End With
                        On Error GoTo 0
                    End If
                Else
                End If
            Next i
        End If

        ' Select the first row if available
        If .lstDependents.ListCount > 0 Then
            .lstDependents.ListIndex = 0
        Else
        End If
        
        ' Convert to tree view with expand/collapse functionality
        .ConvertToTreeView
        .Show vbModeless
    End With
End Sub

Private Function GetCellValueAsString(cell As Range) As String
    On Error Resume Next
    If IsError(cell.value) Then
        GetCellValueAsString = "#ERROR"
    ElseIf IsEmpty(cell.value) Then
        GetCellValueAsString = ""
    Else
        GetCellValueAsString = CStr(cell.value)
    End If
    On Error GoTo 0
End Function


Private Function GetReferencesRecursive(aCell As Range, blnPrecedents As Boolean, currentLevel As Integer, maxLevel As Integer, processedCells As Collection) As String
    
    ' Check if we've reached maximum depth
    If currentLevel > maxLevel Then
        Exit Function
    End If
    
    ' Check for circular references
    Dim cellKey As String
    cellKey = aCell.address(External:=True)
    
    On Error Resume Next
    Dim temp As String
    temp = processedCells(cellKey)
    Dim circularError As Long
    circularError = Err.Number
    
    If circularError = 0 Then
        ' Already processed this cell - circular reference detected
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    
    ' Add this cell to processed collection
    processedCells.Add cellKey, cellKey
    
    ' Store original selection to restore later
    Dim originalSelection As Range
    Set originalSelection = Selection
    
    ' Collect ALL precedents for this cell using a safer approach
    Dim precedentsList As New Collection
    Dim allResults As String
    
    ' Use a safer method to get precedents
    Dim precedentsStr As String
    precedentsStr = GetDirectReferences(aCell, blnPrecedents)
    
    ' Parse the precedents string and add to collection
    If Len(precedentsStr) > 0 Then
        Dim precedentAddresses() As String
        precedentAddresses = Split(precedentsStr, vbCr)
        
        Dim addr As String
        Dim i As Long
        For i = 0 To UBound(precedentAddresses)
            addr = Trim(precedentAddresses(i))
            
            If Len(addr) > 0 Then
                ' Try to get the range from the address
                On Error Resume Next
                Dim precedentCell As Range
                Set precedentCell = Range(addr)
                Dim rangeError As Long
                rangeError = Err.Number
                
                If rangeError = 0 And Not precedentCell Is Nothing Then
                    
                    ' Check for duplicates
                    Dim isDuplicate As Boolean
                    isDuplicate = False
                    Dim j As Long
                    For j = 1 To precedentsList.Count
                        If precedentsList(j).address(External:=True) = precedentCell.address(External:=True) Then
                            isDuplicate = True
                            Exit For
                        End If
                    Next j
                    
                    If Not isDuplicate Then
                        precedentsList.Add precedentCell
                    Else
                    End If
                Else
                End If
                On Error GoTo 0
            Else
            End If
        Next i
    Else
        
    End If
    
    ' Restore original selection
    originalSelection.Select
    
    ' Now process each precedent: display it and immediately recurse (proper tree structure)
    Dim precedent As Range
    For Each precedent In precedentsList
        ' Display this precedent
        allResults = allResults & String(currentLevel * 2, " ") & "L" & currentLevel & ": " & precedent.address(External:=True) & vbCr
        
        ' Check if this precedent is a range (multiple cells)
        If precedent.Cells.Count > 1 Then
            ' This is a range - expand it to show individual cells
            
            Dim cell As Range
            For Each cell In precedent.Cells
                ' Only process if this individual cell hasn't been processed yet
                Dim individualCellKey As String
                individualCellKey = cell.address(External:=True)
                
                On Error Resume Next
                Dim tempCheck As String
                tempCheck = processedCells(individualCellKey)
                If Err.Number <> 0 Then
                    ' Cell hasn't been processed yet
                    On Error GoTo 0
                    
                    ' Display each cell in the range as a sub-item
                    allResults = allResults & String((currentLevel + 1) * 2, " ") & "L" & (currentLevel + 1) & ": " & cell.address(External:=True) & " (from range " & precedent.address(External:=True) & ")" & vbCr
                    
                    ' Recursively trace precedents of this individual cell
                    allResults = allResults & GetReferencesRecursive(cell, blnPrecedents, currentLevel + 2, maxLevel, processedCells)
                Else
                    ' Cell already processed - skip to avoid duplicates
                    On Error GoTo 0
                End If
            Next cell
        Else
            ' Single cell - recurse normally
            allResults = allResults & GetReferencesRecursive(precedent, blnPrecedents, currentLevel + 1, maxLevel, processedCells)
        End If
    Next precedent
    
    GetReferencesRecursive = allResults
End Function

' Helper function to safely get direct precedents without navigation issues
Private Function GetDirectReferences(aCell As Range, blnPrecedents As Boolean) As String
    
    ' Early exit for dependents if cell is unlikely to have dependents
    If Not blnPrecedents Then
        ' For dependents: only proceed if cell has a non-empty value
        If IsEmpty(aCell.Value) Then
            GetDirectReferences = ""
            Exit Function
        End If
    End If
    
    Dim originalSelection As Range
    Dim originalSheet As Worksheet
    Set originalSelection = Selection
    Set originalSheet = ActiveSheet
    
    ' Safely select the target cell (handle cross-sheet references)
    On Error Resume Next
    aCell.Parent.Activate  ' Activate the worksheet first
    Dim activateError As Long
    activateError = Err.Number
    
    If activateError <> 0 Then
        ' If we can't activate the sheet, return empty result
        On Error GoTo 0
        GetDirectReferences = ""
        Exit Function
    End If
    
    aCell.Select  ' Now select the cell
    Dim selectError As Long
    selectError = Err.Number
    
    If selectError <> 0 Then
        ' If we can't select the cell, return empty result
        On Error GoTo 0
        originalSheet.Activate
        originalSelection.Select
        GetDirectReferences = ""
        Exit Function
    End If
    On Error GoTo 0
    
    Dim i As Long
    Dim results As String
    Dim safetyLimit As Long
    safetyLimit = Val(GetSetting("XLerate", "FormulaTracing", "SafetyLimit", "100"))
    
    aCell.Parent.ClearArrows
    
    ' Check if cell has references before calling Show methods to avoid beep
    If blnPrecedents Then
        ' For precedents: only call if cell has a formula
        If aCell.HasFormula Then
            aCell.ShowPrecedents
        End If
    Else
        ' For dependents: only proceed if cell has a value that could be referenced
        ' Similar to HasFormula check for precedents
        If Not IsEmpty(aCell.Value) Then
            ' Cell has a value, so it might have dependents - call ShowDependents to create arrows
            aCell.ShowDependents
        End If
    End If

    ' Collect direct precedents/dependents with better error handling
    i = 0
    Do
        i = i + 1
        
        On Error Resume Next
        Dim navResult As Range
        Set navResult = aCell.NavigateArrow(blnPrecedents, 1, i)
        Dim navError As Long
        navError = Err.Number
        
        ' Check if navigation was successful and we didn't return to original cell
        If navError = 0 And Not navResult Is Nothing Then
            
            If navResult.address(External:=True) <> aCell.address(External:=True) Then
                results = results & navResult.address(External:=True) & vbCr
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
        On Error GoTo 0
        
        ' Safety check to prevent infinite loops
        If i > safetyLimit Then
            Exit Do
        End If
        
    Loop

    ' Collect precedents from other sheets/workbooks
    i = 1
    Do
        i = i + 1
        
        On Error Resume Next
        Set navResult = aCell.NavigateArrow(blnPrecedents, i, 1)
        Dim crossNavError As Long
        crossNavError = Err.Number
        
        If crossNavError = 0 And Not navResult Is Nothing Then
            
            If navResult.address(External:=True) <> aCell.address(External:=True) Then
                results = results & navResult.address(External:=True) & vbCr
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
        On Error GoTo 0
        
        ' Safety check
        If i > safetyLimit Then
            Exit Do
        End If
        
    Loop

    aCell.Parent.ClearArrows
    
    ' Restore original worksheet and selection
    originalSheet.Activate
    originalSelection.Select
    
    GetDirectReferences = results
End Function
