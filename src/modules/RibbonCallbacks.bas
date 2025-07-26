Attribute VB_Name = "RibbonCallbacks"
Option Explicit

'Callback for customUI.onLoad
Public myRibbon As IRibbonUI

'Store ribbon reference
Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    Set myRibbon = ribbon
End Sub

'Callback for Trace Precedents button
Public Sub FindAndDisplayPrecedents(control As IRibbonControl)
    On Error Resume Next
    Application.Run "ShowTracePrecedents"
    On Error GoTo 0
End Sub

'Callback for Trace Dependents button
Public Sub FindAndDisplayDependents(control As IRibbonControl)
    On Error Resume Next
    Application.Run "ShowTraceDependents"
    On Error GoTo 0
End Sub

'Callback for Horizontal Formula Consistency button
Public Sub OnCheckHorizontalConsistency(control As IRibbonControl)
    On Error Resume Next
    Application.Run "CheckHorizontalConsistency"
    On Error GoTo 0
End Sub

Public Sub DoCycleNumberFormat(control As IRibbonControl)
     On Error Resume Next
    Application.Run "ModNumberFormat.CycleNumberFormat"
    On Error GoTo 0
End Sub

Public Sub DoCycleCellFormat(control As IRibbonControl)
    On Error Resume Next
    Application.Run "ModCellFormat.CycleCellFormat"
    On Error GoTo 0
End Sub

Public Sub DoCycleDateFormat(control As IRibbonControl)
    On Error Resume Next
    Application.Run "ModDateFormat.CycleDateFormat"
    On Error GoTo 0
End Sub

Public Sub DoCycleTextStyle(control As IRibbonControl)
    On Error Resume Next
    Application.Run "ModTextStyle.CycleTextStyle"
    On Error GoTo 0
End Sub

Public Sub ShowSettingsForm(control As IRibbonControl)
    Debug.Print "ShowSettingsForm callback was triggered"
    ShowSettings   ' Direct call instead of Application.Run
End Sub

Public Sub DoAutoColorCells(control As IRibbonControl)
    Debug.Print "DoAutoColorCells callback started"
    AutoColorCells control
    Debug.Print "DoAutoColorCells callback ended"
End Sub