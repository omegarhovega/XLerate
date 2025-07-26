Option Explicit

' TreeNode structure for hierarchical display
Private Type TreeNode
    text As String
    level As Integer
    IsExpanded As Boolean
    hasChildren As Boolean
    IsVisible As Boolean
    originalData As String  ' Store original listbox data (address, value, formula)
End Type

Private TreeData() As TreeNode
Private TreeCount As Integer
Private OriginalListData As Collection  ' Store original 3-column data

' Resizer functionality
Private FormResizer As clResizer
Private MinWidth As Single
Private MinHeight As Single
Private LastWidth As Single
Private LastHeight As Single

' Form-level dimension constants
Private Const FORM_PADDING As Long = 5        ' Consistent padding around all edges

' Column width constants
Private Const COLUMN_WIDTH_ADDRESS As Long = 125
Private Const COLUMN_WIDTH_VALUE As Long = 50
Private Const COLUMN_WIDTH_FORMULA As Long = 125

' Control spacing constants
Private Const CONTROL_SPACING As Long = 15     ' Vertical space between controls
Private Const HEADER_OFFSET As Long = 12       ' Space between headers and listbox
Private Const SCROLLBAR_WIDTH As Long = 20     ' Width of vertical scrollbar

Private Const INNER_WIDTH As Long = COLUMN_WIDTH_ADDRESS + COLUMN_WIDTH_VALUE + COLUMN_WIDTH_FORMULA + SCROLLBAR_WIDTH  ' Width of content
Private Const FORM_WIDTH As Long = INNER_WIDTH + (4 * FORM_PADDING)  ' Total form width including padding
Private Const FORM_HEIGHT As Long = 250       ' Total form height

' Formula box constants
Private Const FORMULA_HEIGHT As Long = 20
Private Const FORMULA_TOP As Long = FORM_PADDING
Private Const FORMULA_WIDTH As Long = INNER_WIDTH  ' Width matches the listbox content

' ListBox constants
Private Const LISTBOX_TOP As Long = FORMULA_TOP + FORMULA_HEIGHT + CONTROL_SPACING
Private Const LISTBOX_WIDTH As Long = INNER_WIDTH
Private Const LISTBOX_HEIGHT As Long = FORM_HEIGHT - LISTBOX_TOP - FORM_PADDING  ' Account for bottom padding


' Function to get column widths string
Public Function GetColumnWidths() As String
    GetColumnWidths = COLUMN_WIDTH_ADDRESS & ";" & _
                     COLUMN_WIDTH_VALUE & ";" & _
                     COLUMN_WIDTH_FORMULA
End Function

Public Sub AddHeaders()
    With lstPrecedents
        ' Add headers using a Label control for each column
        Dim headerLabel1 As MSForms.Label
        Set headerLabel1 = Me.Controls.Add("Forms.Label.1", "lblHeader1")
        With headerLabel1
            .Top = lstPrecedents.Top - HEADER_OFFSET
            .Left = lstPrecedents.Left + FORM_PADDING
            .Caption = "Address"
            .Width = COLUMN_WIDTH_ADDRESS
        End With
        
        Dim headerLabel2 As MSForms.Label
        Set headerLabel2 = Me.Controls.Add("Forms.Label.1", "lblHeader2")
        With headerLabel2
            .Top = lstPrecedents.Top - HEADER_OFFSET
            .Left = lstPrecedents.Left + COLUMN_WIDTH_ADDRESS + FORM_PADDING
            .Caption = "Value"
            .Width = COLUMN_WIDTH_VALUE
        End With
        
        Dim headerLabel3 As MSForms.Label
        Set headerLabel3 = Me.Controls.Add("Forms.Label.1", "lblHeader3")
        With headerLabel3
            .Top = lstPrecedents.Top - HEADER_OFFSET
            .Left = lstPrecedents.Left + COLUMN_WIDTH_ADDRESS + COLUMN_WIDTH_VALUE + FORM_PADDING
            .Caption = "Formula"
            .Width = COLUMN_WIDTH_FORMULA
        End With
    End With
End Sub

Private Sub UserForm_Initialize()
    
    ' Set form caption and size
    Me.Caption = "Trace Precedents"
    Me.Width = FORM_WIDTH
    Me.Height = FORM_HEIGHT + (4 * FORM_PADDING) - 2
    
    ' Store minimum and current dimensions
    MinWidth = Me.Width
    MinHeight = Me.Height
    LastWidth = Me.Width
    LastHeight = Me.Height
    
    ' Initialize resizer
    Set FormResizer = New clResizer
    FormResizer.NewForm Me, Me.Left, Me.Top, Me.Width, Me.Height, Me.Caption, 100
    FormResizer.MaintainAspectRatio = False  ' Allow independent width/height resize
    FormResizer.Zoomable = False  ' We'll handle our own resizing logic
    
    ' Initialize tree data
    ReDim TreeData(1 To 100)
    TreeCount = 0
    Set OriginalListData = New Collection
    
    ' Initialize the formula text box
    With Me.Controls.Add("Forms.TextBox.1", "txtFormula")
        .Top = FORMULA_TOP
        .Left = FORM_PADDING
        .Width = FORMULA_WIDTH
        .Height = FORMULA_HEIGHT
        .BackColor = RGB(240, 240, 240)
        .Locked = True
        .MultiLine = True
        .Font.Size = 10
    End With
    
    ' Initialize the list box with adjusted positioning
    With lstPrecedents
        .Top = LISTBOX_TOP
        .Left = FORM_PADDING
        .Width = LISTBOX_WIDTH
        .Height = LISTBOX_HEIGHT
        .ColumnCount = 3
        .ColumnWidths = GetColumnWidths()
        .Font.Size = 10
        .Font.Name = "Consolas"
    End With
    
    ' Ensure form dimensions are properly applied before resizing controls
    DoEvents  ' Allow form to fully initialize
    
    ' Now resize controls to match the actual form dimensions
    ResizeControls
End Sub

Private Sub lstPrecedents_Click()
    On Error Resume Next
    If lstPrecedents.ListIndex < 0 Then Exit Sub
    
    Dim selectedText As String
    selectedText = lstPrecedents.List(lstPrecedents.ListIndex, 0)
    
    ' Navigate to the cell - extract address from display text
    NavigateToCell selectedText
    On Error GoTo 0
End Sub


Private Sub lstPrecedents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim selectedText As String
    Dim nodeIndex As Integer
    Dim currentIndex As Integer
    
    If KeyCode = vbKeyEscape Then
        Unload Me
    ElseIf KeyCode = vbKeyReturn Then
        ' Navigate to cell on Enter
        If lstPrecedents.ListIndex >= 0 Then
            selectedText = lstPrecedents.List(lstPrecedents.ListIndex, 0)
            NavigateToCell selectedText
        End If
    ElseIf KeyCode = vbKeyRight Then
        ' Expand node on Right arrow
        If lstPrecedents.ListIndex >= 0 Then
            currentIndex = lstPrecedents.ListIndex
            selectedText = lstPrecedents.List(currentIndex, 0)
            nodeIndex = FindNodeByDisplayText(selectedText)
            
            If nodeIndex > 0 And TreeData(nodeIndex).hasChildren And Not TreeData(nodeIndex).IsExpanded Then
                TreeData(nodeIndex).IsExpanded = True
                RefreshTreeAndRestoreIndex currentIndex
            End If
        End If
    ElseIf KeyCode = vbKeyLeft Then
        ' Collapse node on Left arrow
        If lstPrecedents.ListIndex >= 0 Then
            currentIndex = lstPrecedents.ListIndex
            selectedText = lstPrecedents.List(currentIndex, 0)
            nodeIndex = FindNodeByDisplayText(selectedText)
            
            If nodeIndex > 0 And TreeData(nodeIndex).hasChildren And TreeData(nodeIndex).IsExpanded Then
                TreeData(nodeIndex).IsExpanded = False
                RefreshTreeAndRestoreIndex currentIndex
            End If
        End If
    End If
    ' Up/Down arrows are handled automatically by the ListBox for navigation
End Sub

' Public method to convert existing ListBox data to tree structure
Public Sub ConvertToTreeView()
    ' Store original data before conversion
    StoreOriginalData
    
    ' Convert indented structure to tree nodes
    ParseListBoxToTree
    
    ' Refresh the tree display
    RefreshTree
End Sub

' Store original 3-column data for navigation
Private Sub StoreOriginalData()
    Set OriginalListData = New Collection
    
    Dim i As Integer
    For i = 0 To lstPrecedents.ListCount - 1
        Dim dataItem As String
        dataItem = lstPrecedents.List(i, 0) & "|" & _
                  lstPrecedents.List(i, 1) & "|" & _
                  lstPrecedents.List(i, 2)
        OriginalListData.Add dataItem
    Next i
End Sub

' Parse existing ListBox indented structure into tree nodes
Private Sub ParseListBoxToTree()
    TreeCount = 0
    
    Dim i As Integer
    For i = 0 To lstPrecedents.ListCount - 1
        Dim addressText As String
        addressText = lstPrecedents.List(i, 0)
        
        ' Extract level from indentation
        Dim level As Integer
        level = GetIndentationLevel(addressText)
        
        ' Extract clean text (remove indentation and level markers)
        Dim cleanText As String
        cleanText = GetCleanText(addressText)
        
        ' Check if this node has children (next item has higher level)
        Dim hasChildren As Boolean
        hasChildren = False
        If i < lstPrecedents.ListCount - 1 Then
            Dim nextLevel As Integer
            nextLevel = GetIndentationLevel(lstPrecedents.List(i + 1, 0))
            hasChildren = (nextLevel > level)
        End If
        
        ' Store original data for this node
        Dim originalData As String
        originalData = lstPrecedents.List(i, 0) & "|" & _
                      lstPrecedents.List(i, 1) & "|" & _
                      lstPrecedents.List(i, 2)
        
        ' Auto-expand level 0 (root) and level 1 nodes to show first level of precedents immediately
        Dim autoExpand As Boolean
        autoExpand = (level = 0)
        
        AddTreeNode cleanText, level, autoExpand, hasChildren, originalData
    Next i
End Sub

' Add a tree node
Private Sub AddTreeNode(nodeText As String, level As Integer, expanded As Boolean, hasChildren As Boolean, originalData As String)
    ' Expand array if needed
    If TreeCount >= UBound(TreeData) Then
        ReDim Preserve TreeData(1 To UBound(TreeData) + 50)
    End If
    
    TreeCount = TreeCount + 1
    With TreeData(TreeCount)
        .text = nodeText
        .level = level
        .IsExpanded = expanded
        .hasChildren = hasChildren
        .IsVisible = True
        .originalData = originalData
    End With
End Sub

' Refresh the tree display
Private Sub RefreshTree()
    Dim i As Integer
    Dim displayText As String
    Dim indent As String
    Dim expandSymbol As String
    
    lstPrecedents.Clear
    
    ' Update visibility based on expanded state
    UpdateVisibility
    
    ' Display visible nodes
    For i = 1 To TreeCount
        If TreeData(i).IsVisible Then
            indent = String(TreeData(i).level * 2, " ")
            
            If TreeData(i).hasChildren Then
                If TreeData(i).IsExpanded Then
                    expandSymbol = "[-] "
                Else
                    expandSymbol = "[+] "
                End If
            Else
                expandSymbol = "    "
            End If
            
            displayText = indent & expandSymbol & TreeData(i).text
            
            ' Parse original data to populate 3 columns
            Dim dataParts() As String
            dataParts = Split(TreeData(i).originalData, "|")
            
            lstPrecedents.AddItem
            With lstPrecedents
                .List(.ListCount - 1, 0) = displayText
                If UBound(dataParts) >= 1 Then .List(.ListCount - 1, 1) = dataParts(1)
                If UBound(dataParts) >= 2 Then .List(.ListCount - 1, 2) = dataParts(2)
            End With
        End If
    Next i
End Sub

' Refresh tree and restore selection to specified index position
Private Sub RefreshTreeAndRestoreIndex(targetIndex As Integer)
    ' Refresh the tree display
    RefreshTree
    
    ' Restore selection to the same index position if possible
    If targetIndex >= 0 And targetIndex < lstPrecedents.ListCount Then
        lstPrecedents.ListIndex = targetIndex
    ElseIf lstPrecedents.ListCount > 0 Then
        ' If target index is out of bounds, select the last available item
        lstPrecedents.ListIndex = lstPrecedents.ListCount - 1
    End If
End Sub

' Update visibility based on parent expansion state
Private Sub UpdateVisibility()
    Dim i As Integer, j As Integer
    Dim parentLevel As Integer
    
    ' First, make all root nodes (level 0) visible
    For i = 1 To TreeCount
        If TreeData(i).level = 0 Then
            TreeData(i).IsVisible = True
        Else
            TreeData(i).IsVisible = False
        End If
    Next i
    
    ' Then, make child nodes visible if their parent is expanded
    For i = 1 To TreeCount
        If TreeData(i).level > 0 Then
            parentLevel = TreeData(i).level - 1
            
            ' Look backwards for the parent
            For j = i - 1 To 1 Step -1
                If TreeData(j).level = parentLevel Then
                    If TreeData(j).IsVisible And TreeData(j).IsExpanded Then
                        TreeData(i).IsVisible = True
                    End If
                    Exit For
                ElseIf TreeData(j).level < parentLevel Then
                    Exit For
                End If
            Next j
        End If
    Next i
End Sub

' Find tree node by display text
Private Function FindNodeByDisplayText(displayText As String) As Integer
    Dim i As Integer
    Dim cleanText As String
    
    ' Clean the display text to get the actual address
    cleanText = displayText
    cleanText = Replace(cleanText, "[-] ", "")
    cleanText = Replace(cleanText, "[+] ", "")
    cleanText = Trim(cleanText)
    
    ' Remove leading spaces (indentation)
    Do While Left(cleanText, 1) = " "
        cleanText = Mid(cleanText, 2)
    Loop
    
    For i = 1 To TreeCount
        If TreeData(i).text = cleanText And TreeData(i).IsVisible Then
            FindNodeByDisplayText = i
            Exit Function
        End If
    Next i
    
    FindNodeByDisplayText = 0
End Function

' Get indentation level from text
Private Function GetIndentationLevel(text As String) As Integer
    ' Count leading spaces and convert to level
    Dim spaceCount As Integer
    Dim i As Integer
    
    For i = 1 To Len(text)
        If Mid(text, i, 1) = " " Then
            spaceCount = spaceCount + 1
        Else
            Exit For
        End If
    Next i
    
    ' Convert spaces to level (assuming 2 spaces per level from original format)
    GetIndentationLevel = spaceCount \ 2
End Function

' Get clean text without indentation and level markers
Private Function GetCleanText(text As String) As String
    Dim cleanText As String
    cleanText = text
    
    ' Remove leading spaces
    cleanText = LTrim(cleanText)
    
    ' Remove level markers like "L1: ", "L2: ", etc.
    Dim colonPos As Long
    colonPos = InStr(cleanText, ": ")
    If colonPos > 0 Then
        cleanText = Mid(cleanText, colonPos + 2)
    End If
    
    GetCleanText = cleanText
End Function

' Navigate to cell based on display text
Private Sub NavigateToCell(displayText As String)
    On Error Resume Next
    
    ' Extract the actual address from the display text
    Dim precedentAddress As String
    precedentAddress = displayText
    
    ' Remove tree formatting
    precedentAddress = Replace(precedentAddress, "[-] ", "")
    precedentAddress = Replace(precedentAddress, "[+] ", "")
    precedentAddress = Replace(precedentAddress, "    ", "")  ' Remove spacing for nodes without children
    precedentAddress = Trim(precedentAddress)
    
    ' Remove leading spaces (indentation)
    Do While Left(precedentAddress, 1) = " "
        precedentAddress = Mid(precedentAddress, 2)
    Loop
    
    ' Remove level markers like "L1: ", "L2: ", etc. (but not for root cell)
    Dim colonPos As Long
    colonPos = InStr(precedentAddress, ": ")
    If colonPos > 0 And InStr(precedentAddress, "L") = 1 Then
        ' Only remove level markers that start with "L" (like "L1: ", "L2: ")
        precedentAddress = Mid(precedentAddress, colonPos + 2)
    End If
    
    ' Clean up the address - remove any extra text like "(from range...)"
    Dim parenPos As Long
    parenPos = InStr(precedentAddress, " (from range")
    If parenPos > 0 Then
        precedentAddress = Trim(Left(precedentAddress, parenPos - 1))
    Else
        precedentAddress = Trim(precedentAddress)
    End If
    
    ' Parse out sheet name and cell address
    Dim exclamationPosition As Integer
    exclamationPosition = InStr(precedentAddress, "!")
    
    If exclamationPosition > 0 Then
        Dim sheetName As String
        Dim cellAddress As String
        
        ' Handle external references like [Workbook]Sheet!Address
        If InStr(precedentAddress, "[") > 0 Then
            ' External reference format: [Workbook]Sheet!Address
            sheetName = Mid(precedentAddress, InStrRev(precedentAddress, "]") + 1, exclamationPosition - InStrRev(precedentAddress, "]") - 1)
        Else
            ' Simple format: Sheet!Address
            sheetName = Left(precedentAddress, exclamationPosition - 1)
        End If
        
        cellAddress = Mid(precedentAddress, exclamationPosition + 1)
        
        ' Clean up sheet name
        If Right(sheetName, 1) = "'" Then
            sheetName = Left(sheetName, Len(sheetName) - 1)
        End If
        If Left(sheetName, 1) = "'" Then
            sheetName = Mid(sheetName, 2)
        End If
        
        ' Navigate to the cell
        Worksheets(sheetName).Activate
        With Worksheets(sheetName).Range(cellAddress)
            .Select
            Me.Controls("txtFormula").text = .formula
        End With
    Else
        ' No sheet specified - assume current sheet
        Range(precedentAddress).Select
        Me.Controls("txtFormula").text = Selection.formula
    End If
    
    On Error GoTo 0
End Sub
' Resizer event handlers
Private Sub UserForm_Activate()
    ' Make the form resizable
    FormResizer.Activate
End Sub

Private Sub UserForm_Resize()
    ' Prevent resizing below minimum dimensions
    If Me.Width < MinWidth Then Me.Width = MinWidth
    If Me.Height < MinHeight Then Me.Height = MinHeight
    
    ' Only resize controls if dimensions actually changed
    If Me.Width <> LastWidth Or Me.Height <> LastHeight Then
        ResizeControls
        LastWidth = Me.Width
        LastHeight = Me.Height
    End If
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
    
    ' Clean up the resizer object to prevent crashes
    Set FormResizer = Nothing
    
    ' Clean up collections
    Set OriginalListData = Nothing
    
    On Error GoTo 0
End Sub

' Resize all controls proportionally
Private Sub ResizeControls()
    On Error Resume Next
    
    ' Calculate new dimensions based on form size
    Dim newInnerWidth As Single
    Dim newInnerHeight As Single
    
    newInnerWidth = Me.Width - (4 * FORM_PADDING)
    newInnerHeight = Me.Height - (4 * FORM_PADDING)
    
    ' Resize formula textbox
    With Me.Controls("txtFormula")
        .Width = newInnerWidth
    End With
    
    ' Resize listbox
    With lstPrecedents
        .Width = newInnerWidth
        .Height = newInnerHeight - FORMULA_HEIGHT - CONTROL_SPACING - HEADER_OFFSET
    End With
    
    ' Update column widths proportionally
    UpdateColumnWidths newInnerWidth
    
    ' Reposition headers
    RepositionHeaders
    
    On Error GoTo 0
End Sub

' Update column widths based on new form width
Private Sub UpdateColumnWidths(newInnerWidth As Single)
    On Error Resume Next
    
    ' Calculate proportional widths
    Dim totalOriginalWidth As Single
    Dim addressWidth As Single
    Dim valueWidth As Single
    Dim formulaWidth As Single
    
    totalOriginalWidth = COLUMN_WIDTH_ADDRESS + COLUMN_WIDTH_VALUE + COLUMN_WIDTH_FORMULA
    
    ' Maintain proportions but account for scrollbar
    Dim availableWidth As Single
    availableWidth = newInnerWidth - SCROLLBAR_WIDTH
    
    addressWidth = (COLUMN_WIDTH_ADDRESS / totalOriginalWidth) * availableWidth
    valueWidth = (COLUMN_WIDTH_VALUE / totalOriginalWidth) * availableWidth
    formulaWidth = (COLUMN_WIDTH_FORMULA / totalOriginalWidth) * availableWidth
    
    ' Update listbox column widths
    lstPrecedents.ColumnWidths = Int(addressWidth) & ";" & Int(valueWidth) & ";" & Int(formulaWidth)
    
    On Error GoTo 0
End Sub

' Reposition header labels
Private Sub RepositionHeaders()
    On Error Resume Next
    
    Dim newInnerWidth As Single
    Dim totalOriginalWidth As Single
    Dim addressWidth As Single
    Dim valueWidth As Single
    
    newInnerWidth = Me.Width - (4 * FORM_PADDING)
    totalOriginalWidth = COLUMN_WIDTH_ADDRESS + COLUMN_WIDTH_VALUE + COLUMN_WIDTH_FORMULA
    
    ' Calculate proportional widths
    Dim availableWidth As Single
    availableWidth = newInnerWidth - SCROLLBAR_WIDTH
    
    addressWidth = (COLUMN_WIDTH_ADDRESS / totalOriginalWidth) * availableWidth
    valueWidth = (COLUMN_WIDTH_VALUE / totalOriginalWidth) * availableWidth
    
    ' Reposition headers
    With Me.Controls("lblHeader1")
        .Width = Int(addressWidth)
    End With
    
    With Me.Controls("lblHeader2")
        .Left = lstPrecedents.Left + addressWidth + FORM_PADDING
        .Width = Int(valueWidth)
    End With
    
    With Me.Controls("lblHeader3")
        .Left = lstPrecedents.Left + addressWidth + valueWidth + FORM_PADDING
        .Width = newInnerWidth - addressWidth - valueWidth - SCROLLBAR_WIDTH
    End With
    
    On Error GoTo 0
End Sub

' Load saved dimensions from registry
Private Sub LoadSavedDimensions()
    On Error Resume Next
    
    Dim savedWidth As Single
    Dim savedHeight As Single
    
    ' Try to load from registry (using GetSetting)
    savedWidth = GetSetting("XLerate", "frmPrecedents", "Width", MinWidth)
    savedHeight = GetSetting("XLerate", "frmPrecedents", "Height", MinHeight)
    
    ' Ensure saved dimensions are not smaller than minimum
    If savedWidth >= MinWidth Then Me.Width = savedWidth
    If savedHeight >= MinHeight Then Me.Height = savedHeight
    
    On Error GoTo 0
End Sub

' Save current dimensions to registry
Private Sub SaveCurrentDimensions()
    On Error Resume Next
    
    ' Only save if dimensions are valid to prevent registry issues
    If Me.Width > 0 And Me.Height > 0 Then
        SaveSetting "XLerate", "frmPrecedents", "Width", Me.Width
        SaveSetting "XLerate", "frmPrecedents", "Height", Me.Height
    End If
    
    On Error GoTo 0
End Sub

