' ModDateFormat
Option Explicit

Private FormatList() As clsFormatType

Public Function GetFormatList() As clsFormatType()
    Debug.Print "=== GetFormatList called ==="
    If Not IsArrayInitialized(FormatList) Then
        Debug.Print "FormatList not initialized - checking saved formats"
        If Not LoadFormatsFromWorkbook() Then
            Debug.Print "No saved formats found - initializing defaults"
            InitializeDateFormats
        End If
    Else
        Debug.Print "FormatList already initialized with " & (UBound(FormatList) + 1) & " formats"
    End If
    
    ' Debug: Show what we're returning
    Debug.Print "Returning FormatList with " & (UBound(FormatList) + 1) & " formats:"
    Dim debugIndex As Integer
    For debugIndex = LBound(FormatList) To UBound(FormatList)
        Debug.Print "  [" & debugIndex & "] " & FormatList(debugIndex).Name & " = '" & FormatList(debugIndex).FormatCode & "'"
    Next debugIndex
    
    GetFormatList = FormatList
    Debug.Print "=== GetFormatList END ==="
End Function

Private Function IsArrayInitialized(ByRef arr As Variant) As Boolean
    On Error Resume Next
    IsArrayInitialized = (UBound(arr) >= 0)
    On Error GoTo 0
End Function



' In ModDateFormat.InitializeDateFormats
Public Sub InitializeDateFormats()
    On Error GoTo ErrorHandler
    Debug.Print "=== InitializeDateFormats START ==="
    
    ' First try to load saved formats
    If Not LoadFormatsFromWorkbook() Then
        Debug.Print "No saved formats found - creating defaults"
        ' Create default formats only if no saved formats exist
        Dim formatObj As clsFormatType
        ReDim FormatList(2)
        
        Set formatObj = New clsFormatType
        formatObj.Name = "Year Only"
        formatObj.FormatCode = "yyyy"
        Set FormatList(0) = formatObj
        Debug.Print "Created format 0: " & FormatList(0).Name
        
        Set formatObj = New clsFormatType
        formatObj.Name = "Month Year"
        formatObj.FormatCode = "mmm-yyyy"
        Set FormatList(1) = formatObj
        Debug.Print "Created format 1: " & FormatList(1).Name
        
        Set formatObj = New clsFormatType
        formatObj.Name = "Full Date"
        formatObj.FormatCode = "dd-mmm-yy"
        Set FormatList(2) = formatObj
        Debug.Print "Created format 2: " & FormatList(2).Name
        
        SaveFormatsToWorkbook
    End If
    
    Debug.Print "=== InitializeDateFormats END ==="
    Exit Sub

ErrorHandler:
    Debug.Print "Error in InitializeDateFormats: " & Err.Description
    Resume Next
End Sub

Public Sub AddFormat(newFormat As clsFormatType)
    Debug.Print "=== AddFormat START ==="
    Debug.Print "Adding format: " & newFormat.Name & " with code: '" & newFormat.FormatCode & "'"
    
    ' Show current FormatList size before adding
    If IsArrayInitialized(FormatList) Then
        Debug.Print "Current FormatList size before adding: " & (UBound(FormatList) + 1)
    Else
        Debug.Print "FormatList not initialized - initializing first"
        InitializeDateFormats
    End If
    
    ' Test what Excel actually interprets the format code as by applying it to a temporary cell
    Dim originalFormatCode As String
    Dim excelInterpretedCode As String
    originalFormatCode = newFormat.FormatCode
    
    ' Use a temporary cell to test the format
    Dim tempCell As Range
    Set tempCell = ActiveSheet.Cells(1, 1) ' Use A1 as temporary
    Dim originalCellFormat As String
    originalCellFormat = tempCell.NumberFormat ' Save original format
    
    ' Apply the new format and see what Excel actually stores
    tempCell.NumberFormat = originalFormatCode
    excelInterpretedCode = tempCell.NumberFormat
    
    ' Restore original format
    tempCell.NumberFormat = originalCellFormat
    
    ' Update the format code to what Excel actually interprets
    newFormat.FormatCode = excelInterpretedCode
    
    Debug.Print "Format code normalization:"
    Debug.Print "  User entered: '" & originalFormatCode & "'"
    Debug.Print "  Excel interprets as: '" & excelInterpretedCode & "'"
    If originalFormatCode <> excelInterpretedCode Then
        Debug.Print "  Format code was modified by Excel - storing Excel's version"
    End If
    
    Dim newIndex As Integer
    newIndex = UBound(FormatList) + 1
    Debug.Print "Adding at index: " & newIndex
    ReDim Preserve FormatList(newIndex)
    Set FormatList(newIndex) = newFormat
    
    Debug.Print "New FormatList size after adding: " & (UBound(FormatList) + 1)
    Debug.Print "Calling SaveFormatsToWorkbook..."
    SaveFormatsToWorkbook
    Debug.Print "=== AddFormat END ==="
End Sub

Public Sub RemoveFormat(index As Integer)
    Debug.Print "Removing format at index: " & index
    Dim i As Integer
    For i = index To UBound(FormatList) - 1
        Set FormatList(i) = FormatList(i + 1)
    Next i
    ReDim Preserve FormatList(UBound(FormatList) - 1)
    SaveFormatsToWorkbook
End Sub

Public Sub UpdateFormat(index As Integer, updatedFormat As clsFormatType)
    Debug.Print "=== UpdateFormat START ==="
    Debug.Print "Updating format at index " & index & ": " & updatedFormat.Name & " with code: '" & updatedFormat.FormatCode & "'"
    
    If index >= 0 And index <= UBound(FormatList) Then
        ' Test what Excel actually interprets the format code as by applying it to a temporary cell
        Dim originalFormatCode As String
        Dim excelInterpretedCode As String
        originalFormatCode = updatedFormat.FormatCode
        
        ' Use a temporary cell to test the format
        Dim tempCell As Range
        Set tempCell = ActiveSheet.Cells(1, 1) ' Use A1 as temporary
        Dim originalCellFormat As String
        originalCellFormat = tempCell.NumberFormat ' Save original format
        
        ' Apply the new format and see what Excel actually stores
        tempCell.NumberFormat = originalFormatCode
        excelInterpretedCode = tempCell.NumberFormat
        
        ' Restore original format
        tempCell.NumberFormat = originalCellFormat
        
        ' Update the format code to what Excel actually interprets
        updatedFormat.FormatCode = excelInterpretedCode
        
        Debug.Print "Format code normalization:"
        Debug.Print "  User entered: '" & originalFormatCode & "'"
        Debug.Print "  Excel interprets as: '" & excelInterpretedCode & "'"
        If originalFormatCode <> excelInterpretedCode Then
            Debug.Print "  Format code was modified by Excel - storing Excel's version"
        End If
        
        Set FormatList(index) = updatedFormat
        SaveFormatsToWorkbook
        Debug.Print "Format updated successfully"
    Else
        Debug.Print "Invalid index: " & index
    End If
    Debug.Print "=== UpdateFormat END ==="
End Sub

Public Sub SaveFormatsToWorkbook()
   Debug.Print "=== SaveFormatsToWorkbook START ==="
   
   ' Wrap the delete in its own error handler
   On Error Resume Next
   ThisWorkbook.CustomDocumentProperties("SavedDateFormats").Delete
   If Err.Number <> 0 Then Debug.Print "Error deleting old property: " & Err.Description
   On Error GoTo ErrorHandler
   
   Dim propValue As String, i As Integer
   For i = LBound(FormatList) To UBound(FormatList)
       Debug.Print "Format " & i & ": " & FormatList(i).Name & " | " & FormatList(i).FormatCode
       propValue = propValue & FormatList(i).Name & "|" & FormatList(i).FormatCode & "||"
   Next i
   
   ThisWorkbook.CustomDocumentProperties.Add Name:="SavedDateFormats", _
       LinkToContent:=False, Type:=msoPropertyTypeString, value:=propValue
       
   Debug.Print "Property added successfully"
   ThisWorkbook.Save
   Debug.Print "Workbook saved successfully"
   Debug.Print "=== SaveFormatsToWorkbook END ==="
   Exit Sub

ErrorHandler:
   Debug.Print "Error in SaveFormatsToWorkbook: " & Err.Description
   MsgBox "Error saving formats: " & Err.Description, vbExclamation
   Resume Next
End Sub

Private Function LoadFormatsFromWorkbook() As Boolean
    Debug.Print "=== LoadFormatsFromWorkbook Debug ==="
    On Error Resume Next
    Dim propValue As String
    propValue = ThisWorkbook.CustomDocumentProperties("SavedDateFormats")
    Debug.Print "Loaded propValue: " & propValue
    If Err.Number <> 0 Then
        Debug.Print "Error loading property: " & Err.Description
        LoadFormatsFromWorkbook = False
        Exit Function
    End If
    On Error GoTo 0

    If propValue = "" Then
        Debug.Print "No saved formats found"
        LoadFormatsFromWorkbook = False
        Exit Function
    End If
    
    Dim formatsArray() As String, formatParts() As String
    formatsArray = Split(propValue, "||")
    Debug.Print "Found " & (UBound(formatsArray) - 1) & " format entries"
    
    ReDim FormatList(UBound(formatsArray) - 1)
    Dim i As Integer
    For i = LBound(formatsArray) To UBound(formatsArray) - 1
        If formatsArray(i) <> "" Then
            Debug.Print "Processing format " & i & ": " & formatsArray(i)
            formatParts = Split(formatsArray(i), "|")
            Set FormatList(i) = New clsFormatType
            FormatList(i).Name = formatParts(0)
            FormatList(i).FormatCode = formatParts(1)
            Debug.Print "Successfully loaded format: " & FormatList(i).Name
        End If
    Next i

    Debug.Print "=== LoadFormatsFromWorkbook Completed ==="
    LoadFormatsFromWorkbook = True
End Function

Public Sub CycleDateFormat()
    Debug.Print "=== CycleDateFormat START ==="
    If Selection Is Nothing Then
        Debug.Print "No selection - exiting"
        Exit Sub
    End If
    If TypeName(Selection) <> "Range" Then
        Debug.Print "Selection is not a Range - exiting"
        Exit Sub
    End If
    
    ' Check if FormatList is initialized
    If Not IsArrayInitialized(FormatList) Then
        Debug.Print "FormatList not initialized - initializing now"
        InitializeDateFormats
    End If
    
    ' Additional check after initialization
    If Not IsArrayInitialized(FormatList) Then
        Debug.Print "Failed to initialize FormatList"
        Exit Sub
    End If
    
    ' Debug: Show current FormatList contents
    Debug.Print "Current FormatList contains " & (UBound(FormatList) + 1) & " formats:"
    Dim debugIndex As Integer
    For debugIndex = LBound(FormatList) To UBound(FormatList)
        Debug.Print "  [" & debugIndex & "] " & FormatList(debugIndex).Name & " = '" & FormatList(debugIndex).FormatCode & "'"
    Next debugIndex
    
    Dim currentFormat As String, nextFormat As String
    Dim found As Boolean
    
    ' Get the format of the first cell in the selection
    currentFormat = Selection.Cells(1).NumberFormat
    Debug.Print "Current cell format: '" & currentFormat & "'"
    
    ' If the selection has multiple cells with different formats,
    ' use the first format in our list
    Dim cell As Range
    For Each cell In Selection
        If cell.NumberFormat <> currentFormat Then
            Debug.Print "Multiple formats detected - using first format in list"
            currentFormat = FormatList(0).FormatCode
            found = True
            Exit For
        End If
    Next cell
    
    If Not found Then
        ' Find the next format in the cycle
        Debug.Print "Searching for current format in FormatList..."
        Dim i As Integer
        For i = LBound(FormatList) To UBound(FormatList)
            Debug.Print "  Comparing '" & currentFormat & "' with FormatList[" & i & "] = '" & FormatList(i).FormatCode & "'"
            If FormatList(i).FormatCode = currentFormat Then
                Debug.Print "  MATCH FOUND at index " & i
                If i < UBound(FormatList) Then
                    nextFormat = FormatList(i + 1).FormatCode
                    Debug.Print "  Next format: FormatList[" & (i + 1) & "] = '" & nextFormat & "'"
                Else
                    nextFormat = FormatList(LBound(FormatList)).FormatCode
                    Debug.Print "  Wrapping to first format: FormatList[" & LBound(FormatList) & "] = '" & nextFormat & "'"
                End If
                found = True
                Exit For
            End If
        Next i
    End If
    
    If Not found Then
        nextFormat = FormatList(LBound(FormatList)).FormatCode
        Debug.Print "Current format not found in list - defaulting to first format: '" & nextFormat & "'"
    End If
    
    Debug.Print "Applying format: '" & nextFormat & "'"
    Selection.NumberFormat = nextFormat
    
    ' Debug: Check what Excel actually stored after applying the format
    Dim actualAppliedFormat As String
    actualAppliedFormat = Selection.Cells(1).NumberFormat
    Debug.Print "Format after applying: '" & actualAppliedFormat & "'"
    If actualAppliedFormat <> nextFormat Then
        Debug.Print "WARNING: Excel modified the format code!"
        Debug.Print "  Original: '" & nextFormat & "'"
        Debug.Print "  Modified: '" & actualAppliedFormat & "'"
    End If
    
    Debug.Print "=== CycleDateFormat END ==="
End Sub
