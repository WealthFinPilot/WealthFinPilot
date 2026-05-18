Attribute VB_Name = "AddButton"
Sub btnS()
Dim Btn As Button
ThisWorkbook.Sheets("WB_Data").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
BtnName = Application.Caller
BtnName = Mid(BtnName, 5)

With AddCat
     .Cat2.Caption = BtnName
     .Show
End With
ThisWorkbook.Sheets("WB_Data").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
End Sub

Sub testtable()
Dim tbcat2 As ListObject
ThisWorkbook.Sheets("WB_Data").Unprotect
Set tbcat2 = ThisWorkbook.Sheets("WB_Data").ListObjects("Loisir")
TbRow = tbcat2.ListRows.Count
ThisWorkbook.Sheets("WB_Data").Protect
End Sub

Sub btnA()
ThisWorkbook.Sheets("WB_Data").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
AddAn.Show
On Error Resume Next
Unload (AddAn)
ThisWorkbook.Sheets("WB_Data").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
End Sub
