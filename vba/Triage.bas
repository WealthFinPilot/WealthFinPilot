Attribute VB_Name = "Triage"
Public tb1, tb2, tby, tbm, tbC, tbline, tbh, REF, TbBudget, TbB1, TbB2, tbRefin, tbLmmmm, DTEP As ListObject
Public cat1d As ListObject
Public cat1i As ListObject
Public Max, Mini, X, Code, CodeB1, CodeB2, SCode, Start, Nb As Long
Public FRow, datarray, FRowB1, FRowB2 As Range
Public Year, NY, LastYear, CheckCode, integerm As Integer
Public YRow, YCode As Variant
Public Année, Logx, Cat_1, Monthvalue, Mois, StrCode, monthnumber, REP, SliceY1, SliceY2, STRY, STRLY, Expense, SIncome, monthGet, RefCat1, MonthFR, monthstring As String
Public monthdate As Date
Public slicerCache As slicerCache
Public Cat2, dataArray, CodeArray As Variant
Public Cat_3, Montant, Pathern, currentDate As String
Public Sarray, SarrayB1, SarrayB2, sortColumn As Range
Public A, B, C, Item, Nbf As Integer
Public UserFormInput As InputAsk
Public Max2 As Integer
Public shp As Shape
Public chartObj As ChartObject
Public LibR1, LibR2, LibR3 As String
Sub Tri_Data()

Application.ScreenUpdating = False
NY = 0
Set tbRefin = ThisWorkbook.Sheets("WB_Data").ListObjects("TbREFIncome")
Set tb1 = ThisWorkbook.Sheets("Input").ListObjects("DataTable")
Set tb2 = ThisWorkbook.Sheets("Data Base").ListObjects("Master")
Set tby = ThisWorkbook.Sheets("WB_Data").ListObjects("Year")
Set tbm = ThisWorkbook.Sheets("WB_Data").ListObjects("Month")
Set tbC = ThisWorkbook.Sheets("WB_Data").ListObjects("Cat")
Set cat1d = ThisWorkbook.Sheets("WB_Data").ListObjects("Dépense")
Set cat1i = ThisWorkbook.Sheets("WB_Data").ListObjects("Income")
Set tbline = ThisWorkbook.Sheets("WB_Data").ListObjects("Data_Line")
Set tbh = ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead")
Set TbBudget = ThisWorkbook.Sheets("WB_Data").ListObjects("Budget")
Set TbB1 = ThisWorkbook.Sheets("Budget1").ListObjects("DataBudget1")
Set TbB2 = ThisWorkbook.Sheets("Budget2").ListObjects("DataBudget2")
Set tbLmmmm = ThisWorkbook.Sheets("Library").ListObjects("tbLibrarymm")


'Data entry-------------------------------------------------------------------
Année = tbh.ListColumns("Header1").DataBodyRange(1, 1).value
Mois = tbh.ListColumns("Header2").DataBodyRange(1, 1).value
Cat_1 = tbh.ListColumns("Header3").DataBodyRange(1, 1).value
Cat_2 = tbh.ListColumns("Header4").DataBodyRange(1, 1).value
Cat_3 = tbh.ListColumns("Header5").DataBodyRange(1, 1).value
Logx = tbh.ListColumns("Header6").DataBodyRange(1, 1).value
Montant = tbh.ListColumns("Header7").DataBodyRange(1, 1).value
lastRow = tbh.ListColumns("Max1").DataBodyRange(1, 1).value
Expense = tbRefin.ListColumns("SuperCat").DataBodyRange(1, 1).value
SIncome = tbRefin.ListColumns("SuperCat").DataBodyRange(2, 1).value
Set SarrayB1 = TbB1.ListColumns(Logx).DataBodyRange
Set SarrayB2 = TbB2.ListColumns(Logx).DataBodyRange
Mini = 0
B = 0
Item = 0
Set REF = ThisWorkbook.Sheets("WB_Data").ListObjects("Refund")
'Sort Table Refund before -------------------------------------------------


' Add sorting rule (smallest to largest)
Set sortColumn = REF.ListColumns("Row").DataBodyRange


'Lets begin !!---------------------------------------------------------------------
For X = Start To lastRow

CheckCode = tb1.ListColumns("Check").DataBodyRange(X, 1).value
Select Case CheckCode
       Case Is = 1
       Nb = Nb - 1
       Nbf = Nbf - 1
       Case Is = 3
       tb1.ListColumns("Check").DataBodyRange(X, 1).value = 1
       Nb = Nb - 1
       Nbf = Nbf - 1
       Case Is = 0
       
Select Case tb1.ListColumns(Cat_1).DataBodyRange(X, 1).value
       Case Is = Expense
       RefCat1 = "Dépense"
       Case Is = SIncome
       RefCat1 = "Income"
End Select

'Reset principal data
Year = tb1.ListColumns(Année).DataBodyRange(X, 1).value
Monthvalue = tb1.ListColumns(Mois).DataBodyRange(X, 1).value
Set FRow = tbm.ListColumns("Mois").DataBodyRange.Find(What:=Monthvalue, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                    On Error Resume Next
                        If Not FRow Is Nothing Then
                           YRow = FRow.Row - 1
                           monthGet = tbLmmmm.ListColumns("English").DataBodyRange(YRow, 1).value
                           MonthFR = tbLmmmm.ListColumns("Français").DataBodyRange(YRow, 1).value
                        Else
                        Exit Sub
                        End If
monthdate = VBA.DateValue(YRow & " " & Year)
monthnumber = Format(YRow, "00")
Cat3 = tb1.ListColumns(Cat_3).DataBodyRange(X, 1).value
integerm = 17 + YRow
monthstring = ThisWorkbook.Sheets("Library").Cells(integerm, "R").value

'Reset door----------------------------------------------------------------
A = 0
C = 0
'Check if this is a new year--------------------------------------------------------------------
       On Error Resume Next
       Set FRow = tbline.ListColumns(Année).DataBodyRange.Find(What:=Year, LookIn:=xlValues, LookAt:=xlWhole)
           If Not FRow Is Nothing Then
                  YRow = FRow.Row - tbline.DataBodyRange.Row + 1
           Else
                  NY = 1
                  YRow = tbline.ListRows.Count + 1
                  tbline.ListColumns(Année).DataBodyRange(YRow, 1).value = Year
                  If Year = ThisWorkbook.Sheets("ControlChart").Range("CM1").value + 1 Then
                  tbline.ListColumns("Solde").DataBodyRange(YRow, 1).value = ThisWorkbook.Sheets("ControlChart").Range("CO4").value
                  tbline.ListColumns("Cible").DataBodyRange(YRow, 1).value = ThisWorkbook.Sheets("ControlChart").Range("CN3").value
                  Else
                  LibR1 = ThisWorkbook.Sheets("Library").Range("A17").value
                  LibR2 = ThisWorkbook.Sheets("Library").Range("A18").value
                  MsgBox (LibR1 & Year & LibR2), vbOK
                  End If
           End If

'Check if this is a new Month----------------------------------------------------------------------
    If tbline.ListColumns(MonthFR).DataBodyRange(YRow, 1).value = "" Then
       A = 1
    End If

'Build SCode---------------------------------------------------------------------

        On Error Resume Next
        Set FRow = REF.ListColumns("C224").DataBodyRange.Find(What:=Cat3, LookIn:=xlValues, LookAt:=xlWhole)
                If Not FRow Is Nothing Then
                       Code = FRow.Row - 1
                       StrCode = Format(REF.ListColumns("Row").DataBodyRange(Code, 1).value, "00")
                       SCode = Year & monthnumber & StrCode

                Else
                      'Category not found------------------------------------------------------------------------------
                       Item = Item + 1
                       C = 1
                           If B = 0 Then
                           'Logic door don't see this message again-----------------------------------------------------
                            LibR1 = ThisWorkbook.Sheets("Library").Range("A19").value
                            REP = Year & " " & Monthvalue & " " & Cat3 & LibR1
                            Set UserFormInput = New InputAsk
                            With UserFormInput
                                 .Label1.Caption = REP
                                 .Show
                            End With
                           
                            ' Check the response
                            Select Case UserFormInput.response
                                   Case Is = 1
                                              B = 1
                                              C = 1
                                   Case Is = 2
                                              C = 1
                                   Case Is = 3
                                           'Stop everything !----------------------------------------------------------------------
                                            Nb = X + 1 - Start - Item
                                            Nbf = lastRow + 1 - Start
                                            LibR1 = ThisWorkbook.Sheets("Library").Range("A20").value
                                            LibR2 = ThisWorkbook.Sheets("Library").Range("A21").value
                                            MsgBox (Item & LibR1 & vbCrLf & vbCrLf & Nb & "/" & Nbf & LibR2), vbOKOnly
                                            Unload UserFormInput
                                            Application.ScreenUpdating = True
                                            Exit Sub
                                   Case Else
                                            Nb = X + 1 - Start - Item
                                            Nbf = lastRow + 1 - Start
                                            LibR1 = ThisWorkbook.Sheets("Library").Range("A20").value
                                            LibR2 = ThisWorkbook.Sheets("Library").Range("A21").value
                                            MsgBox (Item & LibR1 & vbCrLf & vbCrLf & Nb & "/" & Nbf & LibR2), vbOKOnly
                                            Unload UserFormInput
                                            Application.ScreenUpdating = True
                                            Exit Sub
                             End Select
                             Unload UserFormInput
                          End If
                End If
     
'Logic Door in order to create template if needed-----------------------------------------------------------------------------

                  If A = 1 And C = 0 Then
                        'Create template with sub template--------------------------------------------------
                        Template
                  ElseIf C = 0 Then
                  Mini = tbline.ListColumns(MonthFR).DataBodyRange(YRow, 1).value
                  Max = Mini + tbline.ListColumns("Row").DataBodyRange(1, 1).value - 1
                  End If
                  
        'If category is valid------------------------------------------------------------------------------------------
        If C = 0 Then
        
                'Check if there is a SCode in tb2--------------------------------------------------------------------

                'Check fast----------------------------------------------------------------------------------------------
                Set Sarray = tb2.ListColumns(Logx).DataBodyRange.Offset(Mini - 1, 0).Resize(Max)
                On Error Resume Next
                Set FRow = Sarray.Find(What:=SCode, After:=tb2.ListColumns(Logx).DataBodyRange(Mini, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows)
                
                If Not FRow Is Nothing Then
                                     Code = FRow.Row - 1
                                    
                                    'Paste Value tb1 to tb2---------------------------------------------------------------------------
                                     tb2.ListColumns(Montant).DataBodyRange(Code, 1).value = tb2.ListColumns(Montant).DataBodyRange(Code, 1).value + tb1.ListColumns(Montant).DataBodyRange(X, 1).value
                                     tb2.ListColumns("Mois (Name)").DataBodyRange(Code, 1).value = monthstring
                                     tb1.ListColumns("Check").DataBodyRange(X, 1).value = 1
                                     
                                     
                Else
                      'Double check (Long)----------------------------------------------------------------------------------
                                    Set Sarray = tb2.ListColumns(Logx).DataBodyRange
                                                On Error Resume Next
                                                Set FRow = dataArray.Find(What:=SCode, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                If Not FRow Is Nothing Then
                                                        Code = FRow.Row - 1
                                                
                                                        'Paste Value tb1 to tb2---------------------------------------------------------------------------
                                                        tb2.ListColumns(Montant).DataBodyRange(Code, 1).value = tb2.ListColumns(Montant).DataBodyRange(Code, 1).value + tb1.ListColumns(Montant).DataBodyRange(X, 1).value
                                                        tb2.ListColumns("Mois (Name)").DataBodyRange(Code, 1).value = monthstring
                                                        tb1.ListColumns("Check").DataBodyRange(X, 1).value = 1

                                                Else

                                                        'Create new line in tb2------------------------------------------------------------------------
                                                        Max2 = tbh.ListColumns("Max2").DataBodyRange(1, 1).value
                                                        tb1.ListColumns(Année).DataBodyRange(X, 1).Resize(, 5).Copy
                                                        tb2.ListColumns(Année).DataBodyRange(Max2 + 1, 1).PasteSpecial Paste:=xlPasteValues
                                                        Application.CutCopyMode = False
                                                        tb2.ListColumns(Mois).DataBodyRange(Max2 + 1, 1).value = monthdate
                                                        tb2.ListColumns(Logx).DataBodyRange(Max2 + 1, 1).value = SCode
                                                        tb2.ListColumns(Cat_1).DataBodyRange(Max2 + 1, 1).value = RefCat1
                                                        tb2.ListColumns("Mois (Name)").DataBodyRange(Max2 + 1, 1).value = monthstring
                                               
                                                        'Paste Value tb1 to tb2---------------------------------------------------------------------------
                                                        tb2.ListColumns(Montant).DataBodyRange(Max2 + 1, 1).value = tb1.ListColumns(Montant).DataBodyRange(X, 1).value
                                                        tb1.ListColumns("Check").DataBodyRange(X, 1).value = 1
                                  

                                                End If
                End If
                
                 'Update Budget1 AND Budget2 if needed--------------------------------------------------------------------------------------------------------
                                                   If TbBudget.ListColumns("Activate").DataBodyRange(1, 1).value = 1 Then
                                                      On Error Resume Next
                                                          Set FRowB1 = SarrayB1.Find(What:=SCode, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                               If Not FRowB1 Is Nothing Then
                                                                CodeB1 = FRowB1.Row - 1
                                                                TbB1.ListColumns(Montant).DataBodyRange(CodeB1, 1).value = TbB1.ListColumns(Montant).DataBodyRange(CodeB1, 1).value + tb1.ListColumns(Montant).DataBodyRange(X, 1).value
                                                                End If
                                                                
                                                    End If
                                                    If TbBudget.ListColumns("Activate").DataBodyRange(2, 1).value = 1 Then
                                                      On Error Resume Next
                                                          Set FRowB2 = SarrayB2.Find(What:=SCode, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                               If Not FRowB2 Is Nothing Then
                                                                CodeB2 = FRowB2.Row - 1
                                                                TbB2.ListColumns(Montant).DataBodyRange(CodeB2, 1).value = TbB2.ListColumns(Montant).DataBodyRange(CodeB2, 1).value + tb1.ListColumns(Montant).DataBodyRange(X, 1).value
                                                                End If
                                                                
                                                    End If
                
                
         End If

End Select
Next X

Nb = X - Start - Item
Nbf = X - Start
If Nb < 0 Then
   Nb = 0
End If
If Nbf < 0 Then
   Nbf = 0
End If
LibR1 = ThisWorkbook.Sheets("Library").Range("A22").value
LibR2 = ThisWorkbook.Sheets("Library").Range("A23").value
LibR3 = ThisWorkbook.Sheets("Library").Range("A21").value
MsgBox (LibR1 & vbCrLf & vbCrLf & Item & LibR2 & vbCrLf & vbCrLf & Nb & "/" & Nbf & LibR3), vbOKOnly


currentDate = Format(Date, "yyyy-mm-dd")
ThisWorkbook.Sheets("Input").Range("M2").value = currentDate


ThisWorkbook.Sheets("Master").ListObjects("Master_1").QueryTable.Refresh BackgroundQuery:=False

    ' Refresh all
    Do While Application.CalculationState <> xlDone
    DoEvents
    Loop

Resort

If NY = 1 Then
ThisWorkbook.RefreshAll
Set slicerCache = ThisWorkbook.SlicerCaches("Slicer_Année")
slicerCache.ClearManualFilter
SliceY1 = "[Master_1].[Année].&[" & Year & "]"
slicerCache.VisibleSlicerItemsList = Array(SliceY1)
Set slicerCache = ThisWorkbook.SlicerCaches("Slicer_Année4")
slicerCache.ClearManualFilter
LastYear = Year - 1
SliceY2 = "[Master_1].[Année].&[" & LastYear & "]"
slicerCache.VisibleSlicerItemsList = Array(SliceY1, SliceY2)
Set slicerCache = ThisWorkbook.SlicerCaches("Slicer_Année2")
slicerCache.ClearManualFilter
slicerCache.VisibleSlicerItemsList = Array(SliceY1, SliceY2)

End If

Set DTEP = ThisWorkbook.Sheets("ControlChart").ListObjects("DTEP")
On Error Resume Next
Set FRow = DTEP.ListColumns("Année").DataBodyRange.Find(What:=Year, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
    If Not FRow Is Nothing Then
    YRow = FRow.Row - 1
    DTEP.ListColumns("TTEP").DataBodyRange(YRow, 1).value = ThisWorkbook.Sheets("ControlChart").Range("FL3").value
    End If
Application.ScreenUpdating = True
End Sub
Sub Template()
Dim i, SCodetemp As Long
Dim StrCodetemp As String

     'Plant Mini new template----------------------------------------------------------------------
      Max = tbline.ListColumns("Max").DataBodyRange(1, 1).value
            If Max = 1 Then
               Max = 0
            End If
      tbline.ListColumns(MonthFR).DataBodyRange(YRow, 1).value = Max + 1
      Mini = Max + 1

      'Plant Template------------------------------------------------------------
      REF.DataBodyRange.Copy
      tb2.ListColumns(Cat_1).DataBodyRange(Mini, 1).PasteSpecial Paste:=xlPasteValues
      Application.CutCopyMode = False

'Get Max new template----------------------------------------------------------------------
      Max = Mini + tbline.ListColumns("Row").DataBodyRange(1, 1).value - 1
      
      'Apply Month, Year and SCode----------------------------------------------------------------
For i = Mini To Max
    StrCodetemp = Format(tb2.ListColumns(Logx).DataBodyRange(i, 1).value, "00")
    SCodetemp = Year & monthnumber & StrCodetemp
    tb2.ListColumns(Année).DataBodyRange(i, 1).value = Year
    tb2.ListColumns(Mois).DataBodyRange(i, 1).value = monthdate
    tb2.ListColumns(Logx).DataBodyRange(i, 1).value = SCodetemp
    tb2.ListColumns("Mois (Name)").DataBodyRange(i, 1).value = monthstring
Next i
tb2.ListColumns(Montant).DataBodyRange(Mini, 1).value = 0.001
End Sub
Sub Resort()
Dim LastCode, GetMonth, LastYear, LastLine, i As Long
Dim LastMonth As String
Dim frenchMonths() As String
Dim rng As Range
Dim dataArr As Variant

    
Set rng = ThisWorkbook.Sheets("Master").ListObjects("Master_1").ListColumns(Logx).DataBodyRange

dataArr = rng.value
    
    ' Loop through the array to process each value
    For i = LBound(dataArr, 1) To UBound(dataArr, 1)
        ' Check if the last two digits of the value are "01"
        If Right(dataArr(i, 1), 2) = "01" Then
            SCode = dataArr(i, 1)
            LastCode = CStr(SCode)
            GetMonth = Mid(SCode, 5, 2)
            GetMonth = GetMonth + 0
            frenchMonths = Split("janvier,février,mars,avril,mai,juin,juillet,août,septembre,octobre,novembre,décembre", ",")
                  If GetMonth >= 1 And GetMonth <= 12 Then
                          LastMonth = frenchMonths(GetMonth - 1)
                  Else
                          Exit Sub
                  End If
           LastYear = CInt(Left(SCode, 4))
           LastLine = ThisWorkbook.Sheets("Master").ListObjects("Master_1").ListColumns(Logx).DataBodyRange(i - 1, 1).Row
           Set FRow = tbline.ListColumns("Année").DataBodyRange.Find(What:=LastYear, LookIn:=xlValues, LookAt:=xlWhole)
           YRow = FRow.Row - tbline.DataBodyRange.Row + 1
           tbline.ListColumns(LastMonth).DataBodyRange(YRow, 1).value = LastLine
        End If
    Next i

End Sub

