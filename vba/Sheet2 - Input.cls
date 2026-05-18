VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public FirstMonth As String
Private Sub AnalyseD_Click()
Dim Sarray, emptyCells As Range

Application.ScreenUpdating = False
ThisWorkbook.Sheets("WB_Data").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value

Set tb1 = ThisWorkbook.Sheets("Input").ListObjects("DataTable")
Set Sarray = tb1.ListColumns("Check").DataBodyRange
On Error Resume Next
Set emptyCells = Sarray.SpecialCells(xlCellTypeBlanks)
On Error GoTo 0

   If Not emptyCells Is Nothing Then
        Set FindFirstEmptyCellInColumn = emptyCells.Cells(1)
            Start = FindFirstEmptyCellInColumn.Row - 1
            Tri_Data
   End If
   
On Error Resume Next

ThisWorkbook.RefreshAll

ThisWorkbook.Sheets("WB_Data").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
Application.ScreenUpdating = True
End Sub

Private Sub Appli_Extrapolation_Click()
Dim result As VbMsgBoxResult
Dim LibR As String
Application.ScreenUpdating = False
ThisWorkbook.Sheets("WB_Data").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Tampon").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
LibR = ThisWorkbook.Sheets("Library").Range("A8").value
result = MsgBox(LibR, vbYesNo)
Select Case result
       Case Is = vbYes
       tampon
       Case Is = vbNo
       ThisWorkbook.Sheets("WB_Data").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
       ThisWorkbook.Sheets("Input").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
       ThisWorkbook.Sheets("Data Base").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
       ThisWorkbook.Sheets("Tampon").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
       Application.ScreenUpdating = True
       Exit Sub
End Select

On Error Resume Next

ThisWorkbook.RefreshAll

ThisWorkbook.Sheets("WB_Data").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Tampon").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
Application.ScreenUpdating = True
End Sub


Private Sub CommandButton1_Click()
Dim tbAn As ListObject
Dim result As VbMsgBoxResult
Dim An As Integer
Dim ANRow As Long
Dim Mois As String
Dim FRow As Range
Dim LibR As String
Application.ScreenUpdating = False
ThisWorkbook.Sheets("WB_Data").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value


Set tbAn = ThisWorkbook.Sheets("WB_Data").ListObjects("Year")
An = VBA.DateTime.Year(Date)

                Set FRow = tbAn.ListColumns("Année").DataBodyRange.Find(What:=An, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                 On Error Resume Next
                 If Not FRow Is Nothing Then
                 Else
                 LibR = ThisWorkbook.Sheets("Library").Range("A9").value
                 result = MsgBox(LibR, vbYesNo)
                 Select Case result
                        Case Is = vbYes
                        ANRow = tbAn.ListRows.Count
                        tbAn.ListColumns("Année").DataBodyRange(ANRow + 1, 1).value = An
                        'Sort Table tbAn before -------------------------------------------------
                        Set sortColumn = tbAn.ListColumns("Année").Range
                        tbAn.Sort.SortFields.Clear
                        tbAn.Sort.Apply
                        tbAn.Sort.SortFields.Add Key:=sortColumn, _
                        SortOn:=xlSortOnValues, Order:=xlAscending, _
                        DataOption:=xlSortNormal
                          With tbAn.Sort
                          .Header = xlYes
                          .MatchCase = False
                          .Apply
                         End With
                        
                        Case Is = vbNo
                        ThisWorkbook.Sheets("WB_Data").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
                        ThisWorkbook.Sheets("Input").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
                        ThisWorkbook.Sheets("Data Base").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
                        Application.ScreenUpdating = True
                        Exit Sub
                End Select
                End If

InputData.Show

ThisWorkbook.Sheets("WB_Data").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
Application.ScreenUpdating = True
End Sub

Private Sub Reseting_Click()
Dim result As VbMsgBoxResult
Dim newDataBodyRange As Range
Dim LibR As String
LibR = ThisWorkbook.Sheets("Library").Range("A10").value
result = MsgBox(LibR, vbYesNo)
       
       If result = vbYes Then
            Application.ScreenUpdating = True
            ThisWorkbook.Sheets("Input").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
            Set tb1 = ThisWorkbook.Sheets("Input").ListObjects("DataTable")

            tb1.DataBodyRange.Clear
            tb1.ListColumns("Année").DataBodyRange(0 + 1, 1).value = "Exemple"
            tb1.ListColumns("Année").DataBodyRange(1, 1).Font.color = RGB(255, 255, 255)
            Set newDataBodyRange = tb1.ListRows(1).Range
            tb1.Resize tb1.Range.Resize(newDataBodyRange.Rows.Count + 1, newDataBodyRange.Columns.Count)
            ThisWorkbook.Sheets("Input").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
            Application.ScreenUpdating = True
        End If
End Sub
