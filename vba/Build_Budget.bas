Attribute VB_Name = "Build_Budget"
Public TbB, tbTP, tbh, REF, tbEx, tb2 As ListObject
Public wsB, wsAB As Worksheet
Public Budget, ABName, TbBName, Logx, MontantE, Montant, Année As String

Sub Budgetisation(BDG)
Dim rng As Range
Dim dataArr As Variant
Dim i As Long
Dim pt As PivotTable
Application.ScreenUpdating = False
ThisWorkbook.Sheets(Tamp).Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
Set tbTP = ThisWorkbook.Sheets(Tamp).ListObjects("Tampon" & BDG)
Set tbh = ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead")
Set tbcat12 = ThisWorkbook.Sheets("WB_Data").ListObjects("Dépense")
Set tbWbB = ThisWorkbook.Sheets("WB_Data").ListObjects("Budget")
Set REF = ThisWorkbook.Sheets("WB_Data").ListObjects("Refund")
Set tbEx = ThisWorkbook.Sheets(Tamp).ListObjects("Extrap" & BDG)
StartE = 1
Année = tbh.ListColumns("Header1").DataBodyRange(1, 1).value
EndE = tbTP.ListColumns("EndE").DataBodyRange(1, 1).value
Logx = tbh.ListColumns("Header6").DataBodyRange(1, 1).value
MontantE = tbh.ListColumns("Header9").DataBodyRange(1, 1).value
Montant = tbh.ListColumns("Header7").DataBodyRange(1, 1).value


'Budget Choice ----------------------------------------------------------------

              TbBName = "Data" & BDG
              ABName = "A_" & BDG
           
 'Create Budget table data set ----------------------------------------------------------------
Set wsB = ThisWorkbook.Sheets(BDG)
Set wsAB = ThisWorkbook.Sheets(ABName)
Set TbB = ThisWorkbook.Sheets(BDG).ListObjects(TbBName)
wsB.Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value

TbB.DataBodyRange.ClearContents
TbB.ListColumns(Année).DataBodyRange(0 + 1, 1).value = "Exemple"
Set newDataBodyRange = TbB.ListRows(1).Range
TbB.Resize TbB.Range.Resize(newDataBodyRange.Rows.Count + 1, newDataBodyRange.Columns.Count)
tbEx.ListColumns(Année).DataBodyRange(StartE, 1).Resize(EndE, 7).Copy
TbB.ListColumns(Année).DataBodyRange(1, 1).PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False
tbEx.ListColumns(MontantE).DataBodyRange.Copy
TbB.ListColumns(MontantE).DataBodyRange(1, 1).PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False
tbEx.ListColumns(Montant).DataBodyRange.Copy
TbB.ListColumns("Last_Montant").DataBodyRange(1, 1).PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False
TbB.ListColumns(MontantE).DataBodyRange.NumberFormat = "General"
TbB.ListColumns(Logx).DataBodyRange.NumberFormat = "General"
TbB.ListColumns(Montant).DataBodyRange.value = 0.001


Set tb2 = ThisWorkbook.Sheets("Data Base").ListObjects("Master")
SCode = TbB.ListColumns(Logx).DataBodyRange(1, 1).value
 Set FRow = tb2.ListColumns(Logx).DataBodyRange.Find(What:=SCode, LookIn:=xlValues, LookAt:=xlWhole)
                        On Error Resume Next
                        If Not FRow Is Nothing Then
                           SRow = FRow.Row - 1
                        Else
                        ' Loop through all pivot tables in worksheet
                        For Each pt In wsAB.PivotTables
                        pt.RefreshTable
                        Next pt
                        ThisWorkbook.Sheets(Tamp).Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
                        wsB.Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
                        Exit Sub
                        End If

'Feed Table Global and Table Dépense data----------------------------------------------------------------
Set rng = TbB.ListColumns(Logx).DataBodyRange

dataArr = rng.value
    ' Loop through the array to process each value
    For i = LBound(dataArr, 1) To UBound(dataArr, 1)
    
    SCode = TbB.ListColumns(Logx).DataBodyRange(i, 1).value
    Set FRow = tb2.ListColumns(Logx).DataBodyRange.Find(What:=SCode, LookIn:=xlValues, LookAt:=xlWhole)
                        On Error Resume Next
                        If Not FRow Is Nothing Then
                           SRow = FRow.Row - 1
                        If SCode = tb2.ListColumns(Logx).DataBodyRange(SRow, 1).value Then
                            TbB.ListColumns(Montant).DataBodyRange(i, 1).value = tb2.ListColumns(Montant).DataBodyRange(SRow, 1).value
                        End If
                        End If
        
    Next i

        ' Loop through all pivot tables in worksheet
        For Each pt In wsAB.PivotTables
            pt.RefreshTable
        Next pt

ThisWorkbook.Sheets(Tamp).Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
wsB.Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
End Sub
