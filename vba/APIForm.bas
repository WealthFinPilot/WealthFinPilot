Attribute VB_Name = "APIForm"
#If VBA7 Then
    Private Declare PtrSafe Function FindWindowA Lib "user32" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetDC Lib "user32" ( _
        ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" ( _
        ByVal hWnd As LongPtr, _
        ByVal hDC As LongPtr) As Long
    Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" ( _
        ByVal crColor As Long) As LongPtr
    Private Declare PtrSafe Function FrameRect Lib "user32" ( _
        ByVal hDC As LongPtr, _
        ByRef lpRect As rect, _
        ByVal hBrush As LongPtr) As Long
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" ( _
        ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" ( _
        ByVal hDC As LongPtr, _
        ByVal nIndex As Long) As Long
#Else
    Private Declare Function FindWindowA Lib "user32" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
    Private Declare Function GetDC Lib "user32" ( _
        ByVal hWnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal hDC As Long) As Long
    Private Declare Function CreateSolidBrush Lib "gdi32" ( _
        ByVal crColor As Long) As Long
    Private Declare Function FrameRect Lib "user32" ( _
        ByVal hDC As Long, _
        ByRef lpRect As RECT, _
        ByVal hBrush As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" ( _
        ByVal hObject As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" ( _
        ByVal hDC As Long, _
        ByVal nIndex As Long) As Long
#End If

Private Type rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Const LOGPIXELSX = 88
Const LOGPIXELSY = 90
Const TWIPSPERINCH = 1440

Sub SetUserFormBorderColor(frm As Object, color As Long)
    Dim hWnd As LongPtr
    Dim hDC As LongPtr
    Dim hBrush As LongPtr
    Dim rect As rect
    Dim dpiX As Long
    Dim dpiY As Long
    Dim twipsPerPixelX As Double
    Dim twipsPerPixelY As Double

    ' Get the handle of the UserForm window
    hWnd = FindWindowA("ThunderDFrame", frm.Caption)
    
    If hWnd = 0 Then Exit Sub
    
    ' Get the device context of the UserForm window
    hDC = GetDC(hWnd)
    
    ' Get the screen DPI settings
    dpiX = GetDeviceCaps(hDC, LOGPIXELSX)
    dpiY = GetDeviceCaps(hDC, LOGPIXELSY)
    
    ' Calculate Twips per Pixel
    twipsPerPixelX = TWIPSPERINCH / dpiX
    twipsPerPixelY = TWIPSPERINCH / dpiY
    
    ' Create a solid brush with the specified color
    hBrush = CreateSolidBrush(color)
    
    ' Get the dimensions of the UserForm window
    With rect
        .Left = 0
        .Top = 0
        .Right = frm.Width * twipsPerPixelX
        .Bottom = frm.Height * twipsPerPixelY
    End With
    
    ' Draw the border
    FrameRect hDC, rect, hBrush
    
    ' Clean up
    DeleteObject hBrush
    ReleaseDC hWnd, hDC
End Sub

Sub HideLine()
Dim ws As Worksheet
    Dim lastRow As Long
ThisWorkbook.Sheets("Input").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
    ' Set the worksheet to the active sheet or specify the sheet name
    Set ws = ThisWorkbook.Sheets("Input")
    
    ' Define the last possible row in the worksheet
    lastRow = ws.Rows.Count
    
    ' Hide rows from 6 to the last possible row
    
     lastCol = ws.Columns.Count
    
    ' Hide columns from 6 (F) to the last possible column
    ws.Columns("Q:" & Split(ws.Cells(1, lastCol).Address, "$")(1)).Hidden = True
End Sub
Sub HideLine1()
Dim ws As Worksheet
    Dim lastRow As Long

    ' Set the worksheet to the active sheet or specify the sheet name
    Set ws = ThisWorkbook.Sheets("DashBoard")
    
    ' Define the last possible row in the worksheet
ws.Rows.Hidden = False
    
    ' Optionally, unhide all rows in the entire sheet
    ' ws.Rows.Hidden = False

End Sub

