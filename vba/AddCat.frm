VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddCat 
   Caption         =   "Ajouter une sous catégorie."
   ClientHeight    =   3040
   ClientLeft      =   112
   ClientTop       =   448
   ClientWidth     =   4576
   OleObjectBlob   =   "AddCat.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "AddCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BtnName As String

Private Sub Cat3_Change()
Me.Cat3.BorderStyle = fmBorderStyleNone
End Sub

Private Sub CommandButton1_Click()
Dim tbcat12, tbcat11, tbcat2, REF As ListObject
Dim Tbrow1, Tbrow2, i As Long
Dim A As Integer
Dim dataArr1, dataArr2 As Variant
Dim Cat10 As String
Dim rng11, rng12 As Range
Application.ScreenUpdating = False

If Me.Cat3 = "" Then
Me.Cat3.BorderStyle = fmBorderStyleSingle
Me.Cat3.BorderColor = vbRed
Else
BtnName = Me.Cat2.Caption
Set tbcat11 = ThisWorkbook.Sheets("WB_Data").ListObjects("Income")
Set tbcat12 = ThisWorkbook.Sheets("WB_Data").ListObjects("Dépense")
Set tbcat2 = ThisWorkbook.Sheets("WB_Data").ListObjects(BtnName)
Set REF = ThisWorkbook.Sheets("WB_Data").ListObjects("Refund")
i = 0
A = 0
Set rng11 = tbcat11.ListColumns(1).DataBodyRange
Set rng12 = tbcat12.ListColumns(1).DataBodyRange

dataArr1 = rng11.value
dataArr2 = rng12.value
    
    ' Loop through the array to process each value
    For i = LBound(dataArr2, 1) To UBound(dataArr2, 1)
            If BtnName = dataArr2(i, 1) Then
               A = 1
               Exit For
            End If
    Next i
    i = 0
    
    If A = 1 Then
    Cat10 = "Dépense"
    Else
            For i = LBound(dataArr1, 1) To UBound(dataArr1, 1)
                    If BtnName = dataArr1(i, 1) Then
                       Cat10 = "Income"
                       Exit For
                    End If
            Next i
    End If

Tbrow1 = tbcat2.ListRows.Count + 1
tbcat2.ListColumns(1).DataBodyRange(Tbrow1, 1).value = Me.Cat3.value
REF.ListRows.Add
REF.Resize REF.Range.Resize(REF.ListRows.Count + 1, REF.ListColumns.Count)
Tbrow2 = REF.ListRows.Count
REF.ListColumns("Cat_1").DataBodyRange(Tbrow2, 1).value = Cat10
REF.ListColumns("Cat_2").DataBodyRange(Tbrow2, 1).value = BtnName
REF.ListColumns("C224").DataBodyRange(Tbrow2, 1).value = Me.Cat3
REF.ListColumns("Row").DataBodyRange(Tbrow2, 1).value = Tbrow2

Unload Me
End If
Application.ScreenUpdating = True
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub
Private Sub UserForm_Initialize()
Dim LibR As String
SetUserFormBorderColor Me, RGB(0, 0, 0)
LibR = ThisWorkbook.Sheets("Library").Range("A30").value
Me.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A31").value
Me.Label1.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A32").value
Me.CommandButton1.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A224").value
Me.CommandButton2.Caption = LibR
End Sub
