VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddAn 
   Caption         =   "Ajouter une année"
   ClientHeight    =   3040
   ClientLeft      =   112
   ClientTop       =   448
   ClientWidth     =   4368
   OleObjectBlob   =   "AddAn.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "AddAn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub An_Change()
Me.An.ForeColor = vbWhite
Me.An.BorderStyle = fmBorderStyleNone
End Sub

Private Sub CommandButton1_Click()
Dim tbAn As ListObject
Dim valueLength As Integer
Dim Year As Variant
Dim ANRow As Long
Set tbAn = ThisWorkbook.Sheets("WB_Data").ListObjects("Year")
Year = Me.An.value
If Year <> "" Then
    If IsNumeric(Year) Then
        valueLength = Len(Year)
          If valueLength = 4 Then
             ANRow = tbAn.ListRows.Count
             Set FRow = tbAn.ListColumns("Année").DataBodyRange.Find(What:=Year, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                 On Error Resume Next
                 If Not FRow Is Nothing Then
                    LibR = ThisWorkbook.Sheets("Library").Range("A226").value
                    MsgBox (LibR), vbOKOnly
                    Exit Sub
                 Else
                 tbAn.ListColumns("Année").DataBodyRange(ANRow + 1, 1).value = Year
                 
                 Me.Hide
                 Exit Sub
                 End If
          Else
             Me.An.ForeColor = vbRed
             Exit Sub
          End If
    Else
        Me.An.ForeColor = vbRed
        Exit Sub
    End If
Else
Me.An.BorderStyle = fmBorderStyleSingle
Me.An.BorderColor = vbRed
End If
End Sub
Private Sub UserForm_Initialize()
Dim LibR As String
SetUserFormBorderColor Me, RGB(0, 0, 0)
LibR = ThisWorkbook.Sheets("Library").Range("A25").value
Me.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A26").value
Me.Label1.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A27").value
Me.Label2.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A28").value
Me.CommandButton1.Caption = LibR
End Sub
