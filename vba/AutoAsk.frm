VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AutoAsk 
   Caption         =   "Faites votre choix"
   ClientHeight    =   5420
   ClientLeft      =   112
   ClientTop       =   448
   ClientWidth     =   5760
   OleObjectBlob   =   "AutoAsk.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AutoAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tbCCY As ListObject

Private Sub BoxBank_Change()
Dim rng As Range
Select Case Me.BoxBank.value
       Case Is = "Bank1"
           Set rng = tbCCY.ListColumns("Module").DataBodyRange.Rows(1).Resize(2)
           Me.LFolder.RowSource = "ModuleChoice"
           tbCCY.ListColumns("SelectBank").DataBodyRange(1, 1).value = "Bank1"
       Case Is = "Bank2"
           Me.LFolder.RowSource = ""
           Me.LFolder.value = tbCCY.ListColumns("Module").DataBodyRange(1, 1).value
           tbCCY.ListColumns("SelectBank").DataBodyRange(1, 1).value = "Bank2"
       Case Is = "Bank3"
           Me.LFolder.RowSource = "ModuleChoice"
           tbCCY.ListColumns("SelectBank").DataBodyRange(1, 1).value = "Bank3"
       Case Is = "WealthFinPilot"
           Me.LFolder.RowSource = ""
           Me.LFolder.value = tbCCY.ListColumns("Module").DataBodyRange(3, 1).value
           tbCCY.ListColumns("SelectBank").DataBodyRange(1, 1).value = ""
       Case esle
           Me.LFolder.RowSource = ""
           tbCCY.ListColumns("SelectBank").DataBodyRange(1, 1).value = ""
End Select

End Sub

Private Sub CommandButton1_Click()
Select Case Me.LFolder.value
       Case Is = tbCCY.ListColumns("Module").DataBodyRange(1, 1).value
       tbCCY.ListColumns("Path").DataBodyRange(3, 1).value = 1
       tbCCY.ListColumns("Path").DataBodyRange(4, 1).value = 1
       Case Is = tbCCY.ListColumns("Module").DataBodyRange(2, 1).value
       tbCCY.ListColumns("Path").DataBodyRange(3, 1).value = 2
       tbCCY.ListColumns("Path").DataBodyRange(4, 1).value = 1
       Case Is = tbCCY.ListColumns("Module").DataBodyRange(3, 1).value
       tbCCY.ListColumns("Path").DataBodyRange(3, 1).value = 3
       tbCCY.ListColumns("Path").DataBodyRange(4, 1).value = 1
       Case Else
       Me.LFolder.BorderStyle = fmBorderStyleSingle
       Me.LFolder.BorderColor = vbRed
End Select
Unload Me
End Sub

Private Sub Commandquitt_Click()
tbCCY.ListColumns("Path").DataBodyRange(4, 1).value = ""
Unload Me
End Sub

Private Sub CommandReset_Click()
tbCCY.ListColumns("Path").DataBodyRange(4, 1).value = 3
tbCCY.ListColumns("Path").DataBodyRange(3, 1).value = ""
Unload Me
End Sub

Private Sub Continue1_Click()
tbCCY.ListColumns("Path").DataBodyRange(4, 1).value = 2
Unload Me
End Sub

Private Sub LFolder_Change()
Me.LFolder.BorderStyle = fmBorderStyleNone
End Sub

Private Sub UserForm_Initialize()
Dim LibR As String
Set tbCCY = ThisWorkbook.Sheets("WB_Data").ListObjects("tbCCY")
LibR = ThisWorkbook.Sheets("Library").Range("A221").value
Me.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A222").value
Me.Label1.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A223").value
Me.Continue1.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A224").value
Me.Commandquitt.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A225").value
Me.CommandButton1.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A115").value
Me.CommandReset.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A251").value
Me.LabelBank.Caption = LibR
If tbCCY.ListColumns("Path").DataBodyRange(3, 1).value = "" Then
Me.Continue1.Enabled = False
End If
End Sub
