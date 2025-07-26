VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clResizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Windows API declarations for 64-bit and 32-bit compatibility
#If VBA7 Then
    Private Declare PtrSafe Function GetActiveWindow Lib "user32.dll" () As LongPtr
    
    Private Declare PtrSafe Function GetWindowLongW Lib "user32.dll" Alias "GetWindowLongPtrW" ( _
        ByVal hWnd As LongPtr, _
        ByVal nIndex As Long) _
        As LongPtr
    
    Private Declare PtrSafe Function SetWindowLongW Lib "user32.dll" Alias "SetWindowLongPtrW" ( _
        ByVal hWnd As LongPtr, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As LongPtr) _
        As LongPtr
    
    Private Declare PtrSafe Function FindWindowW Lib "user32" ( _
        ByVal lpClassName As LongPtr, _
        ByVal lpWindowName As LongPtr) _
        As LongPtr
    
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
    
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    
    Private Declare PtrSafe Function ReleaseDC Lib "user32" ( _
        ByVal hWnd As LongPtr, _
        ByVal hDC As LongPtr) _
        As Long
    
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" ( _
        ByVal hDC As LongPtr, _
        ByVal nIndex As Long) _
        As Long
#Else
    Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
    
    Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, _
        ByVal nIndex As Long) _
        As Long
    
    Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) _
        As Long
    
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) _
        As Long
    
    Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
    
    Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    
    Private Declare Function ReleaseDC Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal hDC As Long) _
        As Long
    
    Private Declare Function GetDeviceCaps Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nIndex As Long) _
        As Long
#End If

' Constants for Windows API
Private Const GWL_STYLE As Long = -16
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_MINIMIZEBOX As Long = &H20000

' Screen metrics constants
Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

' Class variables
Public Zoomable As Boolean
Public MaintainAspectRatio As Boolean
Private OriginalLeft As Single
Private OriginalTop As Single
Private OriginalWidth As Single
Private OriginalHeight As Single
Private OriginalCaption As String
Private OriginalZoom As Integer
Private FormHandle As LongPtr

' Initialize the resizer
Public Sub NewForm(frm As Object, Left As Single, Top As Single, Width As Single, Height As Single, Caption As String, Zoom As Integer)
    OriginalLeft = Left
    OriginalTop = Top
    OriginalWidth = Width
    OriginalHeight = Height
    OriginalCaption = Caption
    OriginalZoom = Zoom
    Zoomable = True
    MaintainAspectRatio = True
End Sub

' Make the form resizable and set Unicode caption
Public Sub Activate(Optional UniCaption As String = "")
    Dim hWnd As LongPtr
    Dim lStyle As LongPtr
    
    ' Find the form window handle
    If UniCaption <> "" Then
        hWnd = FindFormHandle(UniCaption)
        If hWnd <> 0 Then
            SetFormCaption hWnd, UniCaption
        End If
    Else
        hWnd = FindFormHandle(OriginalCaption)
    End If
    
    If hWnd <> 0 Then
        FormHandle = hWnd
        
        ' Get current window style
        #If VBA7 Then
            lStyle = GetWindowLongW(hWnd, GWL_STYLE)
        #Else
            lStyle = GetWindowLong(hWnd, GWL_STYLE)
        #End If
        
        ' Add resizable style
        lStyle = lStyle Or WS_THICKFRAME Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX
        
        ' Set new window style
        #If VBA7 Then
            SetWindowLongW hWnd, GWL_STYLE, lStyle
        #Else
            SetWindowLong hWnd, GWL_STYLE, lStyle
        #End If
    End If
End Sub

' Handle form resize events
Public Function FormResized(Left As Single, Top As Single, Width As Single, Height As Single, ByRef NewHeight As Single, ByRef NewZoom As Integer) As Boolean
    If Not Zoomable Then
        FormResized = False
        Exit Function
    End If
    
    ' Calculate zoom based on width change
    Dim WidthRatio As Single
    WidthRatio = Width / OriginalWidth
    
        ' Calculate new height to maintain aspect ratio if required
    If MaintainAspectRatio Then
        NewHeight = OriginalHeight * WidthRatio
    Else
        NewHeight = Height
    End If
    
    ' Calculate new zoom level
    NewZoom = Int(OriginalZoom * WidthRatio)
    If NewZoom < 10 Then NewZoom = 10
    If NewZoom > 400 Then NewZoom = 400
    
    FormResized = True
End Function

' Center form on screen
Public Sub CFCalc(WidthPercent As Single, HeightPercent As Single, Limit2Screen As Boolean, ByRef Left As Single, ByRef Top As Single, ByRef Width As Single, ByRef Height As Single)
    Dim ScreenWidth As Single
    Dim ScreenHeight As Single
    
    ' Get screen dimensions
    ScreenWidth = GetScreenWidth()
    ScreenHeight = GetScreenHeight()
    
    ' Calculate dimensions
    If WidthPercent < 0 Then
        ' Negative values mean percentage of original size
        Width = OriginalWidth * Abs(WidthPercent) / 100
        Height = OriginalHeight * Abs(HeightPercent) / 100
    Else
        ' Positive values mean percentage of screen size
        Width = ScreenWidth * WidthPercent / 100
        Height = ScreenHeight * HeightPercent / 100
    End If
    
    ' Limit to screen if requested
    If Limit2Screen Then
        If Width > ScreenWidth Then Width = ScreenWidth * 0.9
        If Height > ScreenHeight Then Height = ScreenHeight * 0.9
    End If
    
    ' Center on screen
    Left = (ScreenWidth - Width) / 2
    Top = (ScreenHeight - Height) / 2
End Sub

' Move form to specific position
Public Sub MoveForm(WidthPercent As Single, HeightPercent As Single, Limit2Screen As Boolean, ByRef Left As Single, ByRef Top As Single, Width As Single, Height As Single)
    Dim ScreenWidth As Single
    Dim ScreenHeight As Single
    
    ' Get screen dimensions
    ScreenWidth = GetScreenWidth()
    ScreenHeight = GetScreenHeight()
    
    ' Calculate position
    Left = ScreenWidth * WidthPercent / 100
    Top = ScreenHeight * HeightPercent / 100
    
    ' Limit to screen if requested
    If Limit2Screen Then
        If Left + Width > ScreenWidth Then Left = ScreenWidth - Width
        If Top + Height > ScreenHeight Then Top = ScreenHeight - Height
        If Left < 0 Then Left = 0
        If Top < 0 Then Top = 0
    End If
End Sub

' Helper functions
Private Function FindFormHandle(Caption As String) As LongPtr
    #If VBA7 Then
        FindFormHandle = FindWindowW(0, StrPtr(Caption))
    #Else
        FindFormHandle = FindWindow(vbNullString, Caption)
    #End If
End Function

Private Sub SetFormCaption(hWnd As LongPtr, Caption As String)
    ' This would require additional API calls to set Unicode caption
    ' For now, we'll keep the original caption functionality
End Sub

Private Function GetScreenWidth() As Single
    Dim hDC As LongPtr
    Dim PixelsPerInchX As Long
    
    hDC = GetDC(0)
    PixelsPerInchX = GetDeviceCaps(hDC, LOGPIXELSX)
    ReleaseDC 0, hDC
    
    GetScreenWidth = GetSystemMetrics(SM_CXSCREEN) * 72 / PixelsPerInchX
End Function

Private Function GetScreenHeight() As Single
    Dim hDC As LongPtr
    Dim PixelsPerInchY As Long
    
    hDC = GetDC(0)
    PixelsPerInchY = GetDeviceCaps(hDC, LOGPIXELSY)
    ReleaseDC 0, hDC
    
    GetScreenHeight = GetSystemMetrics(SM_CYSCREEN) * 72 / PixelsPerInchY
End Function
