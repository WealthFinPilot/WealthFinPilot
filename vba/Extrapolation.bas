Attribute VB_Name = "Extrapolation"
Public tbTP, tbEx, Extra, tbEM, tbTPC, tbline, tbWbB, tb2, tbh, REF, tbCCY, tbMonth, tbLmm As ListObject
Public NbMonth, LastYear, SearchNBMonth, MRow As Integer
Public StartE, EndE, ERow, RowL As Long
Public LastDate, NewDate, EMD1, EMD2 As Date
Public newDataBodyRange As Range
Public i, B, C, X, MonthSupp As Integer
Public Line As Variant
Public Année, Logx, S_Month, STCode, PlusM, TestCode, Cat, MontantE, Montant, formulaString, EM, Tamp, BDG, CRY, chineseText, LangVal, langCode, SearchfrMonth, ENMonth, FRMonth As String
Public StartRow, CurentRow, YRow As Long
Sub tampon()
Dim LibR1, LibR2 As String
Dim SMArray() As Variant
Dim Sarraymmmm As Range
Dim SarrayB As Range
Application.ScreenUpdating = False

Set tbMonth = ThisWorkbook.Sheets("WB_Data").ListObjects("Month")
Set tbLmm = ThisWorkbook.Sheets("Library").ListObjects("tbLibrarymm")
Set tb2 = ThisWorkbook.Sheets("Master").ListObjects("Master_1")
Set tbline = ThisWorkbook.Sheets("WB_Data").ListObjects("Data_Line")
Set tbWbB = ThisWorkbook.Sheets("WB_Data").ListObjects("Budget")
Set tbh = ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead")
Set REF = ThisWorkbook.Sheets("WB_Data").ListObjects("Refund")
Set tbCCY = ThisWorkbook.Sheets("WB_Data").ListObjects("tbCCY")
Logx = tbh.ListColumns("Header6").DataBodyRange(1, 1).value
Année = tbh.ListColumns("Header1").DataBodyRange(1, 1).value
CRY = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
Mois = tbh.ListColumns("Header2").DataBodyRange(1, 1).value
Montant = tbh.ListColumns("Header7").DataBodyRange(1, 1).value
MontantE = tbh.ListColumns("Header9").DataBodyRange(1, 1).value
Cat = tbh.ListColumns("Header8").DataBodyRange(1, 1).value
NbMonth = tbh.ListColumns("NbMonth").DataBodyRange(1, 1).value
RowL = tbline.ListColumns("Row").DataBodyRange(1, 1).value
X = 0
B = 0
If NbMonth <= 0 Then
   LibR1 = ThisWorkbook.Sheets("Library").Range("A12").value
   LibR2 = ThisWorkbook.Sheets("Library").Range("A13").value
   MsgBox (LibR1 & vbCrLf & vbCrLf & LibR2), vbOKOnly
   Exit Sub
End If


A = 0
ExtraP.An = ""
ExtraP.Mois = ""
ExtraP.Show
On Error Resume Next
A = ExtraP.A
Year = ExtraP.Year
BDG = ExtraP.BDG
Monthvalue = ExtraP.Monthvalue
Unload (ExtraP)

'ensure there is a good response to ExtraP---------------------------------
    If A = 0 Then
       Exit Sub
    End If
    
    
'Setting table for Budget-----------------------------------------------------------
         Tamp = "Tampon" & BDG
         Set tbEx = ThisWorkbook.Sheets(Tamp).ListObjects("Extrap" & BDG)
         Set tbTP = ThisWorkbook.Sheets(Tamp).ListObjects("Tampon" & BDG)
         Set tbTPC = ThisWorkbook.Sheets(Tamp).ListObjects("tbTPC" & BDG)
         Set tbEM = ThisWorkbook.Sheets(Tamp).ListObjects("EMois" & BDG)
         
ThisWorkbook.Sheets(Tamp).Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value

Select Case BDG
     Case "Budget1"
     tbWbB.ListColumns("Input").DataBodyRange(1, 1).value = 1
     tbWbB.ListColumns("Input").DataBodyRange(2, 1).value = 0
     
     Case "Budget2"
     tbWbB.ListColumns("Input").DataBodyRange(1, 1).value = 0
     tbWbB.ListColumns("Input").DataBodyRange(2, 1).value = 1
End Select


Set FRow = tbMonth.ListColumns(Mois).DataBodyRange.Find(What:=Monthvalue, LookIn:=xlValues, LookAt:=xlWhole)
               On Error Resume Next
               If Not FRow Is Nothing Then
                  YRow = FRow.Row - 1
                  ENMonth = tbLmm.ListColumns("English").DataBodyRange(YRow, 1).value
                  FRMonth = tbLmm.ListColumns("Français").DataBodyRange(YRow, 1).value
                  monthdate = VBA.DateValue(YRow & " " & Year)
               Else
                 monthdate = DateValue("1 " & Monthvalue & " " & Year)
               End If



Set tbline = ThisWorkbook.Sheets("WB_Data").ListObjects("Data_Line")
monthnumber = Format(Month(monthdate), "00")
MonthSupp = monthnumber + 1
LastYear = Year - 1
'Check if there is enough month------------------------------------------------------------------------
 Set FRow = tbline.ListColumns(Année).DataBodyRange.Find(What:=LastYear, LookIn:=xlValues, LookAt:=xlWhole)
               On Error Resume Next
               If Not FRow Is Nothing Then
                             YRow = FRow.Row - tbline.DataBodyRange.Row + 1
                             
                             For i = 2 To 13
                                   If tbline.ListColumns(i).DataBodyRange(YRow, 1).value <> "" Then
                                      X = X + 1
                                   End If
                             Next i

                End If
Set FRow = tbline.ListColumns(Année).DataBodyRange.Find(What:=Year, LookIn:=xlValues, LookAt:=xlWhole)
               On Error Resume Next
               If Not FRow Is Nothing Then
                             YRow = FRow.Row - tbline.DataBodyRange.Row + 1
                             For i = 2 To 13
                                   If tbline.ListColumns(i).DataBodyRange(YRow, 1).value <> "" Then
                                      If tbline.ListColumns(i).DataBodyRange(YRow, 1).value = tbline.ListColumns(FRMonth).DataBodyRange(YRow, 1).value Then
                                         Exit For
                                      Else
                                      X = X + 1
                                      End If
                                   End If
                             Next i
                Else
                LibR1 = ThisWorkbook.Sheets("Library").Range("A14").value
                LibR2 = ThisWorkbook.Sheets("Library").Range("A15").value
                MsgBox (LibR1 & vbCrLf & vbCrLf & LibR2), vbOKOnly
                Exit Sub
                End If

If MonthSupp = 13 Then
X = 12
End If
If X <= 12 Then
      Extrap_petit
      Exit Sub
End If


'Built Last year Array------------------------------------------------------
           
           Set FRow = tbline.ListColumns(Année).DataBodyRange.Find(What:=LastYear, LookIn:=xlValues, LookAt:=xlWhole)
               On Error Resume Next
               If Not FRow Is Nothing Then
                             YRow = FRow.Row - tbline.DataBodyRange.Row + 1
                             For i = MonthSupp To 13
                                   If tbline.ListColumns(i).DataBodyRange(YRow, 1).value <> "" Then
                                      LastMonth = i
                                      EndMonth = monthnumber + 0
                                      EndYear = Year + 1
                                      Exit For
                                   End If
                             
                             
                             Next i
                Else
                LibR1 = ThisWorkbook.Sheets("Library").Range("A14").value
                LibR2 = ThisWorkbook.Sheets("Library").Range("A15").value
                MsgBox (LibR1 & vbCrLf & vbCrLf & LibR2), vbOKOnly
                Exit Sub
                End If

    
'Built tbEX with SCode starting and Curent SCode------------------------------------------------------------
S_Month = Format(LastMonth, "00")
STCode = LastYear & S_Month & "01"

           On Error Resume Next
           Set FRow = tb2.ListColumns(Logx).DataBodyRange.Find(What:=STCode, LookIn:=xlValues, LookAt:=xlWhole)

               If Not FRow Is Nothing Then
                      StartRow = FRow.Row - tb2.DataBodyRange.Row + 1
               End If


      PlusoneM = monthnumber + 1
      If PlusoneM > 12 Then
      PlusoneM = 1
      PlusM = Format(PlusoneM, "00")
      TestCode = Year + 1 & PlusM & "01"
      Else
      PlusM = Format(PlusoneM, "00")
      TestCode = Year & PlusM & "01"
      End If

           On Error Resume Next
           Set FRow = tb2.ListColumns(Logx).DataBodyRange.Find(What:=TestCode, LookIn:=xlValues, LookAt:=xlWhole)

               If Not FRow Is Nothing Then
                      CurentRow = FRow.Row - StartRow - 1
               Else
                      CurentRow = tb2.ListRows.Count - StartRow + 1
               End If

'Copy one year of input data tb2 to Extrap---------------------------------------------------------
tbEx.DataBodyRange.ClearContents
tbEx.ListColumns(Année).DataBodyRange(0 + 1, 1).value = "Exemple"
Set newDataBodyRange = tbEx.ListRows(1).Range
tbEx.Resize tbEx.Range.Resize(newDataBodyRange.Rows.Count + 1, newDataBodyRange.Columns.Count)
tb2.ListColumns(Année).DataBodyRange(StartRow, 1).Resize(CurentRow, 7).Copy
tbEx.ListColumns(Année).DataBodyRange(1, 1).PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False


'Copy Montant to Extrap Montant-----------------------------------------------------------------------------
Set newDataBodyRange = tbEx.ListColumns(Montant).DataBodyRange
newDataBodyRange.Copy
tbEx.ListColumns(MontantE).DataBodyRange(1, 1).PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False


'Copy data Ref to tbTP and tbTPC---------------------------------------------------------
tbTP.DataBodyRange.ClearContents
tbTP.ListColumns(Cat).DataBodyRange(0 + 1, 1).value = "Exemple"
Set newDataBodyRange = tbTP.ListRows(1).Range
tbTP.Resize tbTP.Range.Resize(newDataBodyRange.Rows.Count + 1, newDataBodyRange.Columns.Count)
REF.ListColumns("C224").DataBodyRange(1, 1).Resize(REF.ListRows.Count, 2).Copy
tbTP.ListColumns(Cat).DataBodyRange(1, 1).PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False
EndE = tbEx.ListRows.Count
tbTP.ListColumns("StartE").DataBodyRange(1, 1) = 1
tbTP.ListColumns("EndE").DataBodyRange(1, 1) = EndE
tbTPC.DataBodyRange.ClearContents
tbTPC.ListColumns(Cat).DataBodyRange(0 + 1, 1).value = "Exemple"
Set newDataBodyRange = tbTPC.ListRows(1).Range
tbTPC.Resize tbTPC.Range.Resize(newDataBodyRange.Rows.Count + 1, newDataBodyRange.Columns.Count)
REF.ListColumns("C224").DataBodyRange(1, 1).Resize(REF.ListRows.Count, 2).Copy
tbTPC.ListColumns(Cat).DataBodyRange(1, 1).PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False


'Start build EMois table-----------------------------------------------------------------------------
tbEM.DataBodyRange.ClearContents
tbEM.ListColumns(Mois).DataBodyRange(0 + 1, 1).value = "Exemple"
Set newDataBodyRange = tbEM.ListRows(1).Range
tbEM.Resize tbEM.Range.Resize(newDataBodyRange.Rows.Count + 1, newDataBodyRange.Columns.Count)
EMD1 = tbEx.ListColumns(Mois).DataBodyRange(1, 1).value
EMD1 = Format(EMD1, "yyyy-mm-dd")
EM = MonthName(Month(EMD1))
EM = Format(EM, "mmmm")
ERow = tbEM.ListRows.Count
tbEM.ListColumns(Mois).DataBodyRange(ERow, 1).value = EM


'Apply new date in order to extrap the data------------------------------------------------------------------
SCode = tbEx.ListColumns(Logx).DataBodyRange(1, 1).value

For i = 1 To EndE
SCode = tbEx.ListColumns(Logx).DataBodyRange(i, 1).value
EMD2 = tbEx.ListColumns(Mois).DataBodyRange(i, 1).value
EMD2 = Format(EMD2, "yyyy-mmmm-dd")

       If EMD2 <> EMD1 Then
          EM = MonthName(Month(EMD2))
          EM = Format(EM, "mmmm")
          ERow = tbEM.ListRows.Count + 1
          tbEM.ListColumns(Mois).DataBodyRange(ERow, 1).value = EM
          EMD1 = EMD2
       End If

Line = Right(SCode, 2)
Line = Line + 0

       If Mid$(SCode, 5, 2) = "01" Then
          B = 1
       End If
       If B = 1 Then
          SCode = EndYear & Mid$(SCode, 5)
          tbEx.ListColumns(Logx).DataBodyRange(i, 1).value = SCode
          tbEx.ListColumns(Année).DataBodyRange(i, 1).value = EndYear
          tbEM.ListColumns(Année).DataBodyRange(ERow, 1).value = EndYear
          LastDate = tbEx.ListColumns(Mois).DataBodyRange(i, 1).value
          NewDate = DateSerial(EndYear, Month(LastDate), Day(LastDate))
          tbEx.ListColumns(Mois).DataBodyRange(i, 1).value = NewDate
       Else
          SCode = Year & Mid$(SCode, 5)
          tbEx.ListColumns(Logx).DataBodyRange(i, 1).value = SCode
          tbEx.ListColumns(Année).DataBodyRange(i, 1).value = Year
          tbEM.ListColumns(Année).DataBodyRange(ERow, 1).value = Year
          LastDate = tbEx.ListColumns(Mois).DataBodyRange(i, 1).value
          NewDate = DateSerial(Year, Month(LastDate), Day(LastDate))
          tbEx.ListColumns(Mois).DataBodyRange(i, 1).value = NewDate
       End If

'Sum the year for average------------------------------------------------------------------------------------------------------
On Error Resume Next
    Set FRow = tbTP.ListColumns("Row").DataBodyRange.Find(What:=Line, LookIn:=xlValues, LookAt:=xlWhole)
            If Not FRow Is Nothing Then
                   YRow = FRow.Row - tbTP.DataBodyRange.Row + 1
                   tbTP.ListColumns(Montant).DataBodyRange(YRow, 1).value = tbTP.ListColumns(Montant).DataBodyRange(YRow, 1).value + tbEx.ListColumns(Montant).DataBodyRange(i, 1).value
            End If
Next i


tbTP.ListColumns(Montant).DataBodyRange.Copy
tbTPC.ListColumns("Montant").DataBodyRange(1, 1).PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False

'Resize month array-------------------------------------------------------------------------------
Line = tbEM.ListRows.Count
tbTP.ListColumns("NbMonth").DataBodyRange(1, 1).value = 12
tbEM.ListRows.Add

'Apply formula--------------------------------------------------------------------------------------------------------------------

formulaString = "=IFERROR([@[Montant]]/$S$2,"""")"
tbTP.ListColumns("MoyenneNC").DataBodyRange.Formula = formulaString
tbTPC.ListColumns("MoyenneC").DataBodyRange.Formula = formulaString

Select Case BDG
     Case "Budget1"
     formulaString = "=IFERROR([@[Montant]]/VLOOKUP([@[Cat_3]],TamponBudget1,4,FALSE),"""")"
     formulaString1 = "=IFERROR(SUM(ExtrapBudget1[Montant_Extrap])/$S$2,"""")"
     Case "Budget2"
     formulaString = "=IFERROR([@[Montant]]/VLOOKUP([@[Cat_3]],TamponBudget2,4,FALSE),"""")"
     formulaString1 = "=IFERROR(SUM(ExtrapBudget2[Montant_Extrap])/$S$2,"""")"
End Select
tbEx.ListColumns("Ratio").DataBodyRange.Formula = formulaString
tbTP.ListColumns("StartE").DataBodyRange(2, 1).Formula = formulaString1


'Format columns----------------------------------------------------------------------------------------------
tbEx.ListColumns("Ratio").DataBodyRange.NumberFormat = "0.00"
tbEx.ListColumns(Montant).DataBodyRange.NumberFormat = "0.00" & CRY
tbEx.ListColumns("Montant_Extrap").DataBodyRange.NumberFormat = "0.00" & CRY
tbTP.ListColumns(Montant).DataBodyRange.NumberFormat = "0.00" & CRY
tbTPC.ListColumns(Montant).DataBodyRange.NumberFormat = "0.00" & CRY

Set tbMonth = ThisWorkbook.Sheets("WB_Data").ListObjects("Month")
ReDim SMArray(1 To 12)
For i = 1 To 12
    SMArray(i) = tbMonth.ListColumns("Mois").DataBodyRange(i, 1).value
Next i


Set tbLmm = ThisWorkbook.Sheets("Library").ListObjects("tbLibrarymm")
LangVal = tbCCY.ListColumns("ChoiceL").DataBodyRange(1, 1).value

chineseText = ChrW(&H4E2D) & ChrW(&H6587)
Select Case LangVal
        Case "Français", "fr"
            langCode = "[$-40C]mmmm" ' French (janv., févr., mars...)
            Set Sarraymmmm = tbLmm.ListColumns(LangVal).DataBodyRange
        Case "English", "en"
            langCode = "[$-409]mmmm" ' English (Jan, Feb, Mar...)
            Set Sarraymmmm = tbLmm.ListColumns(LangVal).DataBodyRange
        Case "Español", "es"
            langCode = "[$-C0A]mmmm" ' Spanish (ene., feb., mar...)
            Set Sarraymmmm = tbLmm.ListColumns(LangVal).DataBodyRange
        Case chineseText, "zh"
            langCode = "[$-804]mmmm" ' Simplified Chinese (1?, 2?, 3?...)
            Set Sarraymmmm = tbLmm.ListColumns(LangVal).DataBodyRange
        Case "Deutsch"
            langCode = "[$-407]mmmm"
            Set Sarraymmmm = tbLmm.ListColumns(LangVal).DataBodyRange
        Case Else
            MsgBox "Invalid language! Please enter French, English, Spanish, German or Chinese.", vbExclamation
            Exit Sub
    End Select

tbEx.ListColumns("Mois").DataBodyRange.NumberFormat = langCode
Set SarrayB = tbEM.ListColumns("Mois").DataBodyRange
SearchNBMonth = tbEM.ListRows.Count - 1
For i = 1 To SearchNBMonth
    SearchfrMonth = tbEM.ListColumns("Mois").DataBodyRange(i, 1).value
           On Error Resume Next
           Set FRow = tbLmm.ListColumns("English").DataBodyRange.Find(What:=SearchfrMonth, LookIn:=xlValues, LookAt:=xlWhole)

               If Not FRow Is Nothing Then
                      MRow = FRow.Row - 1
               End If
   SarrayB.Replace What:=SarrayB(i), Replacement:=SMArray(MRow), LookAt:=xlWhole
Next i

Call Moyenner(Tamp, BDG)
End Sub

Sub Extrap_petit()
Dim CatchMonth As Integer
Dim LibR1 As String
Dim SMArray() As Variant
Dim Sarraymmmm As Range
Dim SarrayB As Range
EndMonth = monthnumber + 0

'Built Last year Array------------------------------------------------------
           On Error Resume Next
           Set FRow = tbline.ListColumns(Année).DataBodyRange.Find(What:=LastYear, LookIn:=xlValues, LookAt:=xlWhole)
               'If there is a last year-------------------------------------------------------------------------
               If Not FRow Is Nothing Then
                             C = 0
                             YRow = FRow.Row - tbline.DataBodyRange.Row + 1
                             For i = 2 To 13
                                   If tbline.ListColumns(i).DataBodyRange(YRow, 1).value <> "" Then
                                      i = i - 1
                                      LastMonth = i
                                      EndYear = Year + 1
                                      Exit For
                                   End If
                             Next i
                Else
                On Error Resume Next
                'If there is no Last year---------------------------------------------------------------------------------
                Set FRow = tbline.ListColumns(Année).DataBodyRange.Find(What:=Year, LookIn:=xlValues, LookAt:=xlWhole)
                        If Not FRow Is Nothing Then
                             C = 1
                             YRow = FRow.Row - tbline.DataBodyRange.Row + 1
                             For i = 2 To 13
                                   If tbline.ListColumns(i).DataBodyRange(YRow, 1).value <> "" Then
                                      i = i - 1
                                      LastMonth = i
                                      EndYear = Year + 1
                                      Exit For
                                   End If
                             Next i
                         End If
                End If

If EndMonth = 12 Then
   C = 1
   LastMonth = 1
End If

    
'Built tbEX with SCode starting and Curent SCode------------------------------------------------------------
S_Month = Format(LastMonth, "00")
If C = 1 Then
   LastYear = Year
End If
STCode = LastYear & S_Month & "01"

           On Error Resume Next
           Set FRow = tb2.ListColumns(Logx).DataBodyRange.Find(What:=STCode, LookIn:=xlValues, LookAt:=xlWhole)

               If Not FRow Is Nothing Then
                  StartRow = FRow.Row - 1
               Else
                  LibR1 = ThisWorkbook.Sheets("Library").Range("A16").value
                  MsgBox (LibR1), vbOKOnly
                  Exit Sub
               End If

PlusoneM = EndMonth + 1
If PlusoneM = 13 Then
PlusoneM = 1
PlusM = Format(PlusoneM, "00")
TestCode = Year + 1 & PlusM & "01"
Else
PlusM = Format(PlusoneM, "00")
TestCode = Year & PlusM & "01"
End If

           On Error Resume Next
           Set FRow = tb2.ListColumns(Logx).DataBodyRange.Find(What:=TestCode, LookIn:=xlValues, LookAt:=xlWhole)

               If Not FRow Is Nothing Then
                      CurentRow = FRow.Row - StartRow - 1
               Else
                      CurentRow = tb2.ListRows.Count - StartRow + 1
               End If
         
If EndMonth < LastMonth Then
   CatchMonth = 12 - LastMonth + 1 + EndMonth
Else
   CatchMonth = EndMonth - LastMonth + 1
End If
         
'Copy one year of input data tb2 to Extrap---------------------------------------------------------
tbEx.DataBodyRange.ClearContents
tbEx.ListColumns(Année).DataBodyRange(0 + 1, 1).value = "Exemple"
Set newDataBodyRange = tbEx.ListRows(1).Range
tbEx.Resize tbEx.Range.Resize(newDataBodyRange.Rows.Count + 1, newDataBodyRange.Columns.Count)
tb2.ListColumns(Année).DataBodyRange(StartRow, 1).Resize(CurentRow, 7).Copy
tbEx.ListColumns(Année).DataBodyRange(1, 1).PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False

'Copy Montant to Extrap Montant-----------------------------------------------------------------------------
Set newDataBodyRange = tbEx.ListColumns(Montant).DataBodyRange
newDataBodyRange.Copy
tbEx.ListColumns(MontantE).DataBodyRange(1, 1).PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False


'Copy data Ref to tbTP and tbTPC---------------------------------------------------------
tbTP.DataBodyRange.ClearContents
tbTP.ListColumns(Cat).DataBodyRange(0 + 1, 1).value = "Exemple"
Set newDataBodyRange = tbTP.ListRows(1).Range
tbTP.Resize tbTP.Range.Resize(newDataBodyRange.Rows.Count + 1, newDataBodyRange.Columns.Count)
REF.ListColumns("C224").DataBodyRange(1, 1).Resize(REF.ListRows.Count, 2).Copy
tbTP.ListColumns(Cat).DataBodyRange(1, 1).PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False
EndE = tbEx.ListRows.Count
tbTP.ListColumns("StartE").DataBodyRange(1, 1) = 1
tbTP.ListColumns("EndE").DataBodyRange(1, 1) = EndE
tbTPC.DataBodyRange.ClearContents
tbTPC.ListColumns(Cat).DataBodyRange(0 + 1, 1).value = "Exemple"
Set newDataBodyRange = tbTPC.ListRows(1).Range
tbTPC.Resize tbTPC.Range.Resize(newDataBodyRange.Rows.Count + 1, newDataBodyRange.Columns.Count)
REF.ListColumns("C224").DataBodyRange(1, 1).Resize(REF.ListRows.Count, 2).Copy
tbTPC.ListColumns(Cat).DataBodyRange(1, 1).PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False


'Start build EMois table-----------------------------------------------------------------------------
tbEM.DataBodyRange.ClearContents
tbEM.ListColumns(Mois).DataBodyRange(0 + 1, 1).value = "Exemple"
Set newDataBodyRange = tbEM.ListRows(1).Range
tbEM.Resize tbEM.Range.Resize(newDataBodyRange.Rows.Count + 1, newDataBodyRange.Columns.Count)
EMD1 = tbEx.ListColumns(Mois).DataBodyRange(1, 1).value
EMD1 = Format(EMD1, "yyyy-mm-dd")
EM = MonthName(Month(EMD1))
EM = Format(EM, "mmmm")
ERow = tbEM.ListRows.Count
tbEM.ListColumns(Mois).DataBodyRange(ERow, 1).value = EM

'Apply new date in order to extrap the data------------------------------------------------------------------
StartE = 1
SCode = tbEx.ListColumns(Logx).DataBodyRange(StartE, 1).value
If Mid$(SCode, 5, 2) = "01" Then
          B = 2
End If


For i = StartE To EndE
SCode = tbEx.ListColumns(Logx).DataBodyRange(i, 1).value
EMD2 = tbEx.ListColumns(Mois).DataBodyRange(i, 1).value
EMD2 = Format(EMD2, "yyyy-mmmm-dd")

       If EMD2 <> EMD1 Then
          EM = MonthName(Month(EMD2))
          EM = Format(EM, "mmmm")
          ERow = ERow + 1
          tbEM.ListColumns(Mois).DataBodyRange(ERow, 1).value = EM
          EMD1 = EMD2
       End If

Line = Right(SCode, 2)
Line = Line + 0

       If Mid$(SCode, 5, 2) = "01" And B <> 2 Then
          B = 1
       End If
       If B = 1 Then
          SCode = (EndYear + 1) & Mid$(SCode, 5)
          tbEx.ListColumns(Logx).DataBodyRange(i, 1).value = SCode
          tbEx.ListColumns(Année).DataBodyRange(i, 1).value = EndYear + 1
          tbEM.ListColumns(Année).DataBodyRange(ERow, 1).value = EndYear + 1
          LastDate = tbEx.ListColumns(Mois).DataBodyRange(i, 1).value
          NewDate = DateSerial(EndYear + 1, Month(LastDate), Day(LastDate))
          tbEx.ListColumns(Mois).DataBodyRange(i, 1).value = NewDate
       Else
             SCode = (Year + 1) & Mid$(SCode, 5)
             tbEx.ListColumns(Logx).DataBodyRange(i, 1).value = SCode
             tbEx.ListColumns(Année).DataBodyRange(i, 1).value = EndYear
             tbEM.ListColumns(Année).DataBodyRange(ERow, 1).value = EndYear
             LastDate = tbEx.ListColumns(Mois).DataBodyRange(i, 1).value
             NewDate = DateSerial(EndYear, Month(LastDate), Day(LastDate))
             tbEx.ListColumns(Mois).DataBodyRange(i, 1).value = NewDate

       End If

'Sum the year for average------------------------------------------------------------------------------------------------------
On Error Resume Next
    Set FRow = tbTP.ListColumns("Row").DataBodyRange.Find(What:=Line, LookIn:=xlValues, LookAt:=xlWhole)
            If Not FRow Is Nothing Then
                   YRow = FRow.Row - tbTP.DataBodyRange.Row + 1
                   tbTP.ListColumns(Montant).DataBodyRange(YRow, 1).value = tbTP.ListColumns(Montant).DataBodyRange(YRow, 1).value + tbEx.ListColumns(Montant).DataBodyRange(i, 1).value
            End If
Next i

tbTP.ListColumns(Montant).DataBodyRange.Copy
tbTPC.ListColumns("Montant").DataBodyRange(1, 1).PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False

Line = tbEM.ListRows.Count
tbTP.ListColumns("NbMonth").DataBodyRange(1, 1).value = CatchMonth
tbEM.ListRows.Add

'Apply formula--------------------------------------------------------------------------------------------------------------------
formulaString = "=IFERROR([@[Montant]]/$S$2,"""")"
tbTP.ListColumns("MoyenneNC").DataBodyRange.Formula = formulaString
tbTPC.ListColumns("MoyenneC").DataBodyRange.Formula = formulaString

Select Case BDG
     Case "Budget1"
     formulaString = "=IFERROR([@[Montant]]/VLOOKUP([@[Cat_3]],TamponBudget1,4,FALSE),"""")"
     formulaString1 = "=IFERROR(SUM(ExtrapBudget1[Montant_Extrap])/$S$2,"""")"
     Case "Budget2"
     formulaString = "=IFERROR([@[Montant]]/VLOOKUP([@[Cat_3]],TamponBudget2,4,FALSE),"""")"
     formulaString1 = "=IFERROR(SUM(ExtrapBudget2[Montant_Extrap])/$S$2,"""")"
End Select
tbEx.ListColumns("Ratio").DataBodyRange.Formula = formulaString
tbTP.ListColumns("StartE").DataBodyRange(2, 1).Formula = formulaString1


'Format columns----------------------------------------------------------------------------------------------
tbEx.ListColumns("Ratio").DataBodyRange.NumberFormat = "0.00"
tbEx.ListColumns(Montant).DataBodyRange.NumberFormat = "0.00" & CRY
tbEx.ListColumns("Montant_Extrap").DataBodyRange.NumberFormat = "0.00" & CRY
tbTP.ListColumns(Montant).DataBodyRange.NumberFormat = "0.00" & CRY
tbTPC.ListColumns(Montant).DataBodyRange.NumberFormat = "0.00" & CRY

Set tbMonth = ThisWorkbook.Sheets("WB_Data").ListObjects("Month")
ReDim SMArray(1 To 12)
For i = 1 To 12
    SMArray(i) = tbMonth.ListColumns("Mois").DataBodyRange(i, 1).value
Next i


Set tbLmm = ThisWorkbook.Sheets("Library").ListObjects("tbLibrarymm")
LangVal = tbCCY.ListColumns("ChoiceL").DataBodyRange(1, 1).value

chineseText = ChrW(&H4E2D) & ChrW(&H6587)
Select Case LangVal
        Case "Français", "fr"
            langCode = "[$-40C]mmmm" ' French (janv., févr., mars...)
            Set Sarraymmmm = tbLmm.ListColumns(LangVal).DataBodyRange
        Case "English", "en"
            langCode = "[$-409]mmmm" ' English (Jan, Feb, Mar...)
            Set Sarraymmmm = tbLmm.ListColumns(LangVal).DataBodyRange
        Case "Español", "es"
            langCode = "[$-C0A]mmmm" ' Spanish (ene., feb., mar...)
            Set Sarraymmmm = tbLmm.ListColumns(LangVal).DataBodyRange
        Case chineseText, "zh"
            langCode = "[$-804]mmmm" ' Simplified Chinese (1?, 2?, 3?...)
            Set Sarraymmmm = tbLmm.ListColumns(LangVal).DataBodyRange
        Case "Deutsch"
            langCode = "[$-407]mmmm"
            Set Sarraymmmm = tbLmm.ListColumns(LangVal).DataBodyRange
        Case Else
            MsgBox "Invalid language! Please enter French, English, Spanish, German or Chinese.", vbExclamation
            Exit Sub
    End Select

tbEx.ListColumns("Mois").DataBodyRange.NumberFormat = langCode
Set SarrayB = tbEM.ListColumns("Mois").DataBodyRange
SearchNBMonth = tbEM.ListRows.Count - 1
For i = 1 To SearchNBMonth
    SearchfrMonth = tbEM.ListColumns("Mois").DataBodyRange(i, 1).value
           On Error Resume Next
           Set FRow = tbLmm.ListColumns("English").DataBodyRange.Find(What:=SearchfrMonth, LookIn:=xlValues, LookAt:=xlWhole)

               If Not FRow Is Nothing Then
                      MRow = FRow.Row - 1
               End If
   SarrayB.Replace What:=SarrayB(i), Replacement:=SMArray(MRow), LookAt:=xlWhole
Next i

Call Moyenner(Tamp, BDG)
End Sub
Sub Moyenner(Tamp, BDG)
Application.ScreenUpdating = False
ThisWorkbook.Sheets(Tamp).Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value

MOY.Show

ThisWorkbook.Sheets(Tamp).Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
Application.ScreenUpdating = True
End Sub


