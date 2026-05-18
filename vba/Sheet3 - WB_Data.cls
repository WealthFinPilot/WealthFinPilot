VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public BtnName As String
Sub AddButton()
Dim rng As Range
Dim Btn As Button

For Each Btn In ThisWorkbook.Sheets("WB_Data").Buttons
         Check = Btn.Name
Btn.Name = "Btn " & "i1_Income22"
Btn.Caption = "i1_Income22"
Next Btn
Set rng = ThisWorkbook.Sheets("WB_Data").Range("A2")
Set Btn = ThisWorkbook.Sheets("WB_Data").Buttons.Add(rng.Left, rng.Top, rng.Width, rng.Height)
   Button_Name = "Année"
     With Btn
     .OnAction = "btnA"
     .Caption = Button_Name
     .Font.Bold = True
     .Name = "Btn " & Button_Name
     End With
End Sub

Private Sub CommandButton1_Click()
ThisWorkbook.Sheets("WB_Data").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Budget1").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Budget2").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("TamponBudget1").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("TamponBudget2").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ChangeCAT.Show
ThisWorkbook.Sheets("WB_Data").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Budget1").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Budget2").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("TamponBudget1").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("TamponBudget2").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
End Sub
