' frmSettingsManager
Option Explicit
' Control declarations
Private NumbersPanel As MSForms.Frame
Private CellsPanel As MSForms.Frame
Private DatesPanel As MSForms.Frame
Private numberSettings As frmNumberSettings
Private WithEvents lstCategories As MSForms.ListBox
Private AutoColorPanel As MSForms.Frame
Private autoColorSettings As frmAutoColor
Private ErrorPanel As MSForms.Frame
Private TextStylesPanel As MSForms.Frame
Private textStyleSettings As frmTextStyle
Private formulaTracingSettings As frmFormulaTracingSettings
Private FormulaTracingPanel As MSForms.Frame


Private Sub UserForm_Initialize()
    Debug.Print "SettingsManager Initialize started"
    On Error GoTo ErrorHandler
    
    ' Initialize form layout
    InitializeFormLayout
    Debug.Print "Form layout initialized"
    
    ' Create navigation listbox with event handling
    Debug.Print "Creating categories listbox"
    Set lstCategories = Me.Controls.Add("Forms.ListBox.1", "lstCategories")
    With lstCategories
        .Left = 12
        .Top = 12
        .Width = 150
        .Height = 450
    End With
    Debug.Print "Categories list created"
    
    ' Create panels
    InitializePanels
    Debug.Print "Panels initialized"
    
    InitializeHierarchyList
    Debug.Print "Hierarchy list initialized"
    
    Debug.Print "SettingsManager Initialize completed successfully"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in SettingsManager Initialize: " & Err.Description & " (Error " & Err.Number & ")"
    Resume Next
End Sub

Private Sub InitializePanels()
    ' Create Numbers panel frame
    Set NumbersPanel = Me.Controls.Add("Forms.Frame.1", "NumbersPanel")
    With NumbersPanel
        .Left = 170
        .Top = 12
        .Width = 410
        .Height = 450
        .Caption = ""
        .BackColor = RGB(255, 255, 255)
        .Visible = False
    End With
    
    ' Create Cells panel frame
    Set CellsPanel = Me.Controls.Add("Forms.Frame.1", "CellsPanel")
    With CellsPanel
        .Left = 170
        .Top = 12
        .Width = 410
        .Height = 450
        .Caption = ""
        .BackColor = RGB(255, 255, 255)
        .Visible = False
    End With
    
    ' Create Dates panel frame
    Set DatesPanel = Me.Controls.Add("Forms.Frame.1", "DatesPanel")
    With DatesPanel
        .Left = 170
        .Top = 12
        .Width = 410
        .Height = 450
        .Caption = ""
        .BackColor = RGB(255, 255, 255)
        .Visible = False
    End With
    
    ' Create Auto-Color panel frame
    Set AutoColorPanel = Me.Controls.Add("Forms.Frame.1", "AutoColorPanel")
    With AutoColorPanel
        .Left = 170
        .Top = 12
        .Width = 410
        .Height = 450
        .Caption = ""
        .BackColor = RGB(255, 255, 255)
        .Visible = False
    End With
    
    ' Create Error panel frame
    Set ErrorPanel = Me.Controls.Add("Forms.Frame.1", "ErrorPanel")
    With ErrorPanel
        .Left = 170
        .Top = 12
        .Width = 410
        .Height = 450
        .Caption = ""
        .BackColor = RGB(255, 255, 255)
        .Visible = False
    End With
    
    ' Create Text Styles panel frame
    Set TextStylesPanel = Me.Controls.Add("Forms.Frame.1", "TextStylesPanel")
    With TextStylesPanel
        .Left = 170
        .Top = 12
        .Width = 410
        .Height = 450
        .Caption = ""
        .BackColor = RGB(255, 255, 255)
        .Visible = False
    End With
    
    ' Create Formula Tracing panel frame
    Set FormulaTracingPanel = Me.Controls.Add("Forms.Frame.1", "FormulaTracingPanel")
    With FormulaTracingPanel
        .Left = 170
        .Top = 12
        .Width = 410
        .Height = 450
        .Caption = ""
        .BackColor = RGB(255, 255, 255)
        .Visible = False
    End With
    
    ' Initialize all settings within their respective panels
    Debug.Print "Initializing all panels..."
    InitializeNumberSettings NumbersPanel
    InitializeCellSettings CellsPanel
    InitializeDateSettings DatesPanel
    InitializeAutoColorSettings AutoColorPanel
    InitializeErrorSettings ErrorPanel
    InitializeTextStyleSettings TextStylesPanel
    InitializeFormulaTracingSettings FormulaTracingPanel
    Debug.Print "All panels initialized"
End Sub

Private Sub InitializeNumberSettings(parentFrame As MSForms.Frame)
    ' Create a new instance of frmNumberSettings
    Dim numberSettings As New frmNumberSettings
    ' Initialize it within the panel
    numberSettings.InitializeInPanel parentFrame
End Sub

Private Sub InitializeCellSettings(parentFrame As MSForms.Frame)
    ' Create a new instance of frmCellSettings
    Dim cellSettings As New frmCellSettings
    ' Initialize it within the panel
    cellSettings.InitializeInPanel parentFrame
End Sub

Private Sub InitializeDateSettings(parentFrame As MSForms.Frame)
    Debug.Print "Initializing date settings"
    Dim dateSettings As New frmDateSettings
    dateSettings.InitializeInPanel parentFrame
    Debug.Print "Date settings initialized"
End Sub

Private Sub InitializeAutoColorSettings(parentFrame As MSForms.Frame)
    Debug.Print "Initializing auto-color settings"
    Set autoColorSettings = New frmAutoColor
    autoColorSettings.InitializeInPanel parentFrame
    Debug.Print "Auto-color settings initialized"
End Sub

Private Sub InitializeErrorSettings(parentFrame As MSForms.Frame)
    Dim errorSettings As New frmErrorSettings
    errorSettings.InitializeInPanel parentFrame
End Sub

Private Sub InitializeTextStyleSettings(parentFrame As MSForms.Frame)
    On Error GoTo ErrorHandler
    
    Debug.Print vbNewLine & "=== InitializeTextStyleSettings START ==="
    Debug.Print "Creating new frmTextStyle instance"
    Set textStyleSettings = New frmTextStyle
    
    Debug.Print "Initializing text style settings in panel"
    Debug.Print "Parent frame is Nothing: " & (parentFrame Is Nothing)
    If Not parentFrame Is Nothing Then
        Debug.Print "Parent frame name: " & parentFrame.Name
    End If
    
    textStyleSettings.InitializeInPanel parentFrame
    Debug.Print "=== InitializeTextStyleSettings END ==="
    Exit Sub

ErrorHandler:
    Debug.Print "Error in InitializeTextStyleSettings: " & Err.Description & " (Error " & Err.Number & ")"
    Resume Next
End Sub

Private Sub InitializeFormulaTracingSettings(parentFrame As MSForms.Frame)
    Debug.Print "Initializing formula tracing settings"
    Set formulaTracingSettings = New frmFormulaTracingSettings
    formulaTracingSettings.InitializeInPanel parentFrame
    Debug.Print "Formula tracing settings initialized"
End Sub


' Show the requested panel and hide others
Private Sub ShowPanel(panelName As String)
    Debug.Print vbNewLine & "=== ShowPanel called ==="
    Debug.Print "panelName: '" & panelName & "'"
    
    ' Debug panel objects
    Debug.Print "TextStylesPanel Is Nothing: " & (TextStylesPanel Is Nothing)
    If Not TextStylesPanel Is Nothing Then
        Debug.Print "TextStylesPanel.Name: " & TextStylesPanel.Name
    End If
    
    ' Hide all panels first
    NumbersPanel.Visible = False
    DatesPanel.Visible = False
    CellsPanel.Visible = False
    AutoColorPanel.Visible = False
    ErrorPanel.Visible = False
    TextStylesPanel.Visible = False
    FormulaTracingPanel.Visible = False
    
    Select Case panelName
        Case "Numbers"
            NumbersPanel.Visible = True
            Debug.Print "Showing Numbers panel"
        Case "Dates"
            DatesPanel.Visible = True
            Debug.Print "Showing Dates panel"
        Case "Cells"
            CellsPanel.Visible = True
            Debug.Print "Showing Cells panel"
        Case "Auto-Color"
            AutoColorPanel.Visible = True
            Debug.Print "Showing Auto-Color panel"
        Case "Error"
            ErrorPanel.Visible = True
            Debug.Print "Showing Error panel"
        Case "Text Styles"
            TextStylesPanel.Visible = True
            Debug.Print "Showing Text Styles panel"
        Case "Formula Tracing"
            FormulaTracingPanel.Visible = True
            Debug.Print "Showing Formula Tracing panel"
    End Select
    
    DebugPanelState
    Debug.Print "=== ShowPanel completed ==="
End Sub

Private Sub UserForm_Terminate()
    Set numberSettings = Nothing
    Set NumbersPanel = Nothing
    Set CellsPanel = Nothing
    Set DatesPanel = Nothing
    Set AutoColorPanel = Nothing
    Set autoColorSettings = Nothing
    Set ErrorPanel = Nothing
    Set TextStylesPanel = Nothing
    Set textStyleSettings = Nothing
    Set FormulaTracingPanel = Nothing
    Set lstCategories = Nothing
End Sub

' Add this event handler for the listbox
Private Sub lstCategories_Click()
    Debug.Print vbNewLine & "=== lstCategories_Click triggered ==="
    On Error GoTo ErrorHandler
    
    Dim selectedCategory As String
    selectedCategory = Trim(lstCategories.text)
    Debug.Print "Selected category: '" & selectedCategory & "'"
    Debug.Print "Category length: " & Len(selectedCategory)
    Debug.Print "ASCII codes: "
    Dim i As Integer
    For i = 1 To Len(selectedCategory)
        Debug.Print "Position " & i & ": " & Asc(Mid(selectedCategory, i, 1))
    Next i
    
    If lstCategories.List(lstCategories.ListIndex, 1) = "HEADER" Then
        Debug.Print "Header clicked: " & selectedCategory
        ' Handle different headers - prevent selection by redirecting to first item under header
        Select Case selectedCategory
            Case "--- FORMATTING ---"
                Debug.Print "Formatting header clicked, selecting Numbers"
                lstCategories.ListIndex = 1  ' Select Numbers (first item under Formatting)
                ShowPanel "Numbers"
            Case "--- FORMULAS ---"
                Debug.Print "Formulas header clicked, selecting Formula Tracing"
                ' Find Formula Tracing item index
                For i = 0 To lstCategories.ListCount - 1
                    If lstCategories.List(i, 0) = "Formula Tracing" Then
                        lstCategories.ListIndex = i
                        ShowPanel "Formula Tracing"
                        Exit For
                    End If
                Next i
            Case Else
                lstCategories.ListIndex = 1  ' Default to Numbers
                ShowPanel "Numbers"
        End Select
        Exit Sub
    End If
    
    Debug.Print "Processing category selection"
    ' Remove any potential hidden characters
    selectedCategory = Replace(selectedCategory, vbTab, "")
    selectedCategory = Replace(selectedCategory, vbCr, "")
    selectedCategory = Replace(selectedCategory, vbLf, "")
    selectedCategory = Trim(selectedCategory)
    
    Select Case selectedCategory
        Case "Numbers"
            ShowPanel "Numbers"
        Case "Dates"
            ShowPanel "Dates"
        Case "Cells"
            ShowPanel "Cells"
        Case "Auto-Color"
            ShowPanel "Auto-Color"
        Case "Text Styles"
            ShowPanel "Text Styles"
        Case "Error"
            ShowPanel "Error"
        Case "Formula Tracing"
            ShowPanel "Formula Tracing"
        Case Else
            Debug.Print "Unknown category: " & selectedCategory
    End Select
    Debug.Print "=== lstCategories_Click completed ==="
    Exit Sub

ErrorHandler:
    Debug.Print "Error in lstCategories_Click: " & Err.Description & " (Error " & Err.Number & ")"
    Resume Next
End Sub

Private Sub DebugPanelState()
    Debug.Print vbNewLine & "=== Panel State Debug ==="
    Debug.Print "NumbersPanel is Nothing: " & (NumbersPanel Is Nothing)
    If Not NumbersPanel Is Nothing Then Debug.Print "NumbersPanel.Visible: " & NumbersPanel.Visible
    
    Debug.Print "DatesPanel is Nothing: " & (DatesPanel Is Nothing)
    If Not DatesPanel Is Nothing Then Debug.Print "DatesPanel.Visible: " & DatesPanel.Visible
    
    Debug.Print "CellsPanel is Nothing: " & (CellsPanel Is Nothing)
    If Not CellsPanel Is Nothing Then Debug.Print "CellsPanel.Visible: " & CellsPanel.Visible
End Sub

Private Sub InitializeFormLayout()
    Me.BackColor = RGB(255, 255, 255)
    Me.Caption = "Settings"
    Me.Width = 600
    Me.Height = 500
    Debug.Print "Form layout set"
End Sub

Private Sub InitializeHierarchyList()
    Debug.Print "Initializing hierarchy list"
    lstCategories.Clear
    
    With lstCategories
        .AddItem "--- FORMATTING ---"
        .List(.ListCount - 1, 1) = "HEADER"
        .AddItem "Numbers"
        .AddItem "Dates"
        .AddItem "Cells"
        .AddItem "Auto-Color"
        .AddItem "Text Styles"
        .AddItem "Error"
        .AddItem "--- FORMULAS ---"
        .List(.ListCount - 1, 1) = "HEADER"
        .AddItem "Formula Tracing"
        .ListIndex = 1
    End With
    
    ' Make headers bold and non-selectable
    FormatCategoryHeaders
    
    ShowPanel "Numbers"
End Sub

Private Sub FormatCategoryHeaders()
    ' This method would ideally format headers as bold, but VBA ListBox has limited formatting
    ' The bold formatting and non-selectable behavior is handled in the click event
    Debug.Print "Headers formatted (bold styling handled in click event)"
End Sub
