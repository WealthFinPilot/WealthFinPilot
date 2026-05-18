VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputAsk 
   Caption         =   "Oups !"
   ClientHeight    =   3160
   ClientLeft      =   112
   ClientTop       =   448
   ClientWidth     =   4712
   OleObjectBlob   =   "InputAsk.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "InputAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public response As Integer
Private Sub GO_Again_Click()
Dim SelectedValue As Integer
response = 2
Me.Hide
End Sub

Private Sub GO_NotAgain_Click()
response = 1
Me.Hide
End Sub

Private Sub Stopit_Click()
Dim SelectedValue As Integer
response = 3
Me.Hide
End Sub

Private Sub UserForm_Initialize()
Dim LibR As String
SetUserFormBorderColor Me, RGB(0, 0, 0)
LibR = ThisWorkbook.Sheets("Library").Range("A92").value
Me.GO_NotAgain.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A93").value
Me.Go_Again.Caption
LibR = ThisWorkbook.Sheets("Library").Range("A94").value
Me.Stopit.Caption = LibR
End Sub
