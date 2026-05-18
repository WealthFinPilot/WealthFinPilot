VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MOY 
   Caption         =   "Console"
   ClientHeight    =   8160
   ClientLeft      =   112
   ClientTop       =   448
   ClientWidth     =   13848
   OleObjectBlob   =   "MOY.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "MOY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Line, monthnumberD As Integer
Public Extra, tbEx, tbTP, tbEM, tbTPC, tbh, REF, TbB, tbCCY, tbMonth, tbLmmmm As ListObject
Public foundLabel, foundLabel1, foundLabe2 As Control
Public MREFRow, StartE, EndE, SRow, EMend, MRow, YRow, SAVCRow As Long
Public Sarray, rng, cell As Range
Public Starti, Endi, Endir, B, C, P, F, MSend, NBM, AM, PORTE_M, retryCount, maxRetries, SameYear, SpecialM1, SpecialM2 As Integer
Public Montant, Logx, Année, Mois, MontantE, monthnumber, SE, monthStart, monthEnd, SCode, SCodeTEST, Tamp, Cat1, MonthTEST, BDG, CRY, CodeM, monthGet, MyValMS, MyValAv, MyTTB, MyTTDiff, MyTTREF, MyTTP, localDec, altDec, MyValP1, MyValP2, vbaSep As String
Public monthdate1, monthdate2, monthdateTEST As Date
Public ValAv, ValPlus, ValP1, ValP2, Valup, ValName, ValRName, LabelName, m As Variant
Public MREF, TT, MPond, MS, Mend, AverageC, Diff, NewVal, MBAV, TTB, TTDiff, TTREF, TTP, ValMS, ValMSCount As Double
Dim DisableEvents As Boolean

Private Sub Average_Change()
If DisableEvents Then
Exit Sub
End If
Me.Average.BorderStyle = fmBorderStyleNone
Me.Average.ForeColor = vbWhite
DisableEvents = True
Me.Pond.value = ""
DisableEvents = False
Me.Pond.BorderStyle = fmBorderStyleNone
Me.Pond.ForeColor = vbWhite
End Sub

Private Sub Cat_Change()
Me.Cat.BorderStyle = fmBorderStyleNone
     For i = 0 To 11
         Set foundLabel = MOY.Controls(ValName(i))
         foundLabel.Caption = ""
         foundLabel.ForeColor = &H808080
         Set foundLabel = MOY.Controls(ValRName(i))
         foundLabel.Caption = ""
         Set foundLabel = MOY.Controls(LabelName(i))
         foundLabel.ForeColor = &H808080
      Next i


If Me.Cat.value = "" Then
   Me.Pond.value = ""
   Me.LabelDiff.Caption = ""
   Me.LabelTTDiff.Caption = ""
   Me.TBudget.Caption = ""
   Me.TREF.Caption = ""
   Me.BAV.Caption = ""
   Me.REFAV.Caption = ""
   Me.Average.value = ""
   
Else
  
  Me.LabelDiff.Caption = ""

  Cat = Me.Cat.value

    Set FRow = tbTP.ListColumns("Cat").DataBodyRange.Find(What:=Cat, LookIn:=xlValues, LookAt:=xlWhole)
           On Error Resume Next
           If Not FRow Is Nothing Then
                  MREFRow = FRow.Row - 1
                  TTREF = tbTP.ListColumns(Montant).DataBodyRange(MREFRow, 1).value
                  MyTTREF = Replace(Format(Abs(TTREF), "0.00"), ".", Application.International(xlDecimalSeparator))
                  Me.TREF.Caption = MyTTREF & CRY
                  MREF = tbTP.ListColumns("MoyenneNC").DataBodyRange(MREFRow, 1).value
                  MyMREF = Replace(Format(MREF, "0.00"), ".", Application.International(xlDecimalSeparator))
                  Me.REFAV.Caption = MyMREF & CRY
                  TTB = tbTPC.ListColumns(Montant).DataBodyRange(MREFRow, 1).value
                  MyTTB = Replace(Format(Abs(TTB), "0.00"), ".", Application.International(xlDecimalSeparator))
                  Me.TBudget.Caption = MyTTB & CRY
                  MBAV = tbTPC.ListColumns("MoyenneC").DataBodyRange(MREFRow, 1).value
                  MyMBAV = Replace(Format(MBAV, "0.00"), ".", Application.International(xlDecimalSeparator))
                  Me.BAV.Caption = MyMBAV & CRY
                  TTDiff = TTB - TTREF
                  MyTTDiff = Replace(Format(TTDiff, "0.00"), ".", Application.International(xlDecimalSeparator))
                  If TTREF = 0 Or TTREF = "" Then
                  TTP = 100
                  Else
                  TTP = ((TTB / TTREF) - 1) * 100
                  MyTTP = Replace(Format(TTP, "0.00"), ".", Application.International(xlDecimalSeparator))
                  End If
                    If TTDiff < 0 Then
                    LibR = ThisWorkbook.Sheets("Library").Range("A125").value
                    Me.GainTotal.Caption = LibR
                    Me.LabelTTDiff.ForeColor = &H80000004
                    Me.LabelTTDiffP.ForeColor = &H80000004
                    ElseIf TTDiff = 0 Then
                    LibR = ThisWorkbook.Sheets("Library").Range("A118").value
                    Me.GainTotal.Caption = LibR
                    Me.LabelTTDiff.ForeColor = &H808080
                    Me.LabelTTDiffP.ForeColor = &H808080
                    ElseIf TTDiff > 0 Then
                    LibR = ThisWorkbook.Sheets("Library").Range("A118").value
                    Me.GainTotal.Caption = LibR
                    Me.LabelTTDiff.ForeColor = &H80000004
                    Me.LabelTTDiffP.ForeColor = &H80000004
                    End If
                    Me.LabelTTDiff.Caption = MyTTDiff & CRY
                    If TTP < 0 Then
                    Me.LabelTTDiffP.Caption = "(" & MyTTP & "%" & ")"
                    Else
                    Me.LabelTTDiffP.Caption = "(+" & MyTTP & "%" & ")"
                    End If
                  Me.MO_1.value = ""
                  Me.MO_2.value = ""
                  Me.BAV.ForeColor = &H808080
                  Me.TBudget.ForeColor = &H808080
                  
                  Set FRow = tbTP.ListColumns("Cat").DataBodyRange.Find(What:=Cat, LookIn:=xlValues, LookAt:=xlWhole)
                             On Error Resume Next
                             If Not FRow Is Nothing Then
                                SRow = FRow.Row - 1
                                SE = tbTP.ListColumns("Row").DataBodyRange(SRow, 1).value
                                SE = Format(SE, "00")
                             End If
                  Set FRow = tbEx.ListColumns("Cat_3").DataBodyRange.Find(What:=Cat, LookIn:=xlValues, LookAt:=xlWhole)
                             On Error Resume Next
                             If Not FRow Is Nothing Then
                                SAVCRow = FRow.Row - 1
                                            
                                     For i = 0 To NBM
                                         Set foundLabel = MOY.Controls(LabelName(i))
                                         If foundLabel <> "" Then
                                         foundLabel.ForeColor = &H808080
                                         monthnumber = Left(foundLabel.Caption, 2)
                                         monthnumber = Format(monthnumber, "00")
                                         SCode = monthnumber & SE
                                         
                                            Set FRow = tbEx.ListColumns(Logx).DataBodyRange.Find(What:="*" & SCode, After:=tbEx.ListColumns(Logx).DataBodyRange(StartE, 1), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                                                If Not FRow Is Nothing Then
                                                   YRow = FRow.Row - 1
                                                   ValMS = tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value
                                                   MyValMS = Replace(Format(ValMS, "0.00"), ".", Application.International(xlDecimalSeparator))
                                                   ValMS = CDbl(MyValMS)
                                                   Set foundLabel = MOY.Controls(ValName(i))
                                                   
                                                        If ValMS = 0 Then
                                                           foundLabel.Caption = "0.00" & CRY
                                                        Else
                                                           foundLabel.Caption = MyValMS & CRY
                                                        End If
                                                        If tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 1 Then
                                                            Me.foundLabel.ForeColor = &H80000004
                                                            Me.BAV.ForeColor = &H80000004
                                                            Me.TBudget.ForeColor = &H80000004
                                                        End If
                                                        
                                                   ValMS = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value
                                                   MyValMS = Replace(Format(ValMS, "0.00"), ".", Application.International(xlDecimalSeparator))
                                                   ValMS = CDbl(MyValMS)
                                                   Set foundLabel = MOY.Controls(ValRName(i))
                                                   
                                                        If ValMS = 0 Then
                                                           foundLabel.Caption = "0.00" & CRY
                                                        Else
                                                           foundLabel.Caption = MyValMS & CRY
                                                        End If
                                                        
                                                 End If
                                                 End If
                                    Next i
                                                
                             End If

                    
            End If
End If
End Sub

Private Sub CommandM_Click()
Dim LibR As String

P = 0
m = 0
F = 0
PORTE_M = 0
Diff = 0

MO1 = Me.MO_1.value

                    Set FRow = tbMonth.ListColumns("Mois").DataBodyRange.Find(What:=MO1, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                    On Error Resume Next
                        If Not FRow Is Nothing Then
                           SpecialM1 = FRow.Row - 1
                           MO1 = tbLmmmm.ListColumns("English").DataBodyRange(SpecialM1, 1).value
                        End If
                  
MO2 = Me.MO_2.value

                   Set FRow = tbMonth.ListColumns("Mois").DataBodyRange.Find(What:=MO2, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                    On Error Resume Next
                        If Not FRow Is Nothing Then
                           SpecialM2 = FRow.Row - 1
                           MO2 = tbLmmmm.ListColumns("English").DataBodyRange(SpecialM2, 1).value
                        End If

ValAv = Me.Average.value
ValAv = Replace(ValAv, localDec, altDec)
If vbaSep = "." Then
ValAv = Replace(ValAv, altDec, Application.International(xlDecimalSeparator))
End If


    If IsNumeric(ValAv) Then
        ValAv = CDbl(ValAv)
        ValAv = ValAv + 0
    Else
        If Me.Average.value = "" Then
        Me.Average.BorderStyle = fmBorderStyleSingle
        Me.Average.BorderColor = vbRed
        Else
        Me.Average.ForeColor = vbRed
        End If
        ValAv = ""
        F = 2
    End If


If Me.Cat = "" Then
Me.Cat.BorderStyle = fmBorderStyleSingle
Me.Cat.BorderColor = vbRed
   F = 2
End If
If MO1 = "" And MO2 = "" Then
Me.MO_1.BorderStyle = fmBorderStyleSingle
Me.MO_1.BorderColor = vbRed
Me.MO_2.BorderStyle = fmBorderStyleSingle
Me.MO_2.BorderColor = vbRed
   F = 2
End If


Set Sarray = tbEx.ListColumns(Logx).DataBodyRange

If (MO1 = "" And MO2 = "") Or Me.Cat = "" Or ValAv = "" Then
B = 0
    If F = 2 Then
       Exit Sub
    End If
'Enter Average in an Array--------------------------------------------------------------------------------------------
ElseIf MO1 <> "" And MO2 <> "" Then
B = 1
         'Reset period------------------------------------
         For i = 0 To NBM
         Set foundLabel = MOY.Controls(LabelName(i))
         foundLabel.ForeColor = &H808080
         Next i

     monthdate1 = VBA.DateValue(SpecialM1 & " " & 2000)
     monthdate2 = VBA.DateValue(SpecialM2 & " " & 2000)
     Starti = Month(monthdate1)
     Endi = Month(monthdate2)
     
     If Endi < Starti Then
     
            'Changing year--------------------------------------------------------------------------------------------
              For i = Starti To MSend
                    monthnumber = Format(i, "00")
                    SCode = monthnumber & SE
                    retryCount = 0
                    Do
                    On Error Resume Next
                    Set FRow = Sarray.Find(What:="*" & SCode, After:=tbEx.ListColumns(Logx).DataBodyRange(StartE, 1), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                    
                        If Not FRow Is Nothing Then
                           YRow = FRow.Row - 1
                           'CalculDiff---------------------------------------------------------------------------------
                           CalculDiff
                           tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = ValAv
                           ValMS = ValAv
                           MyValMS = Replace(Format(ValAv, "0.00"), ".", Application.International(xlDecimalSeparator))
                           If PORTE_M = 1 Then
                                Set foundLabel1 = MOY.Controls(ValName(m))
                                foundLabel1.Caption = MyValMS & CRY
                                Set foundLabel2 = MOY.Controls(LabelName(m))
                                foundLabel2.ForeColor = &H80000004
                                m = m + 1
                             Else
                                For m = 0 To NBM
                                    Set foundLabel2 = MOY.Controls(LabelName(m))
                                    If Left(foundLabel2.Caption, 2) = monthnumber Then
                                       Set foundLabel1 = MOY.Controls(ValName(m))
                                       foundLabel1.Caption = MyValMS & CRY
                                       foundLabel2.ForeColor = &H80000004
                                       m = Right(foundLabel1.Name, 2)
                                       PORTE_M = 1
                                       Exit For
                                    End If
                                Next m
                            End If
                          'Check changment----------------------------------------------------------------------------------------------------------------
                           If tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value Then
                              tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 0
                              foundLabel1.ForeColor = &H808080
                           Else
                              tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 1
                              foundLabel1.ForeColor = &H80000004
                           End If
                        
                         Exit Do
                         Else
                         'Apply Cat3 missing ----------------------------------------------------------------
                         retryCount = retryCount + 1
                              If retryCount > maxRetries Then
                              Exit Do
                              Else
                                     tbEx.ListRows.Add
                                     REF.ListColumns("Cat_1").DataBodyRange(SE, 1).Resize(1, 3).Copy
                                     tbEx.ListColumns("Cat_1").DataBodyRange(EndE + 1, 1).PasteSpecial Paste:=xlPasteValues
                                     Application.CutCopyMode = False
                                     monthnumberD = monthnumber + 0
                                     MonthTEST = Application.WorksheetFunction.Text(DateSerial(2025, monthnumberD, 1), "[$-409]mmmm")
                                     On Error Resume Next
                                     Set FRow = tbLmmmm.ListColumns("English").DataBodyRange.Find(MonthTEST, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                                     If Not FRow Is Nothing Then
                                     YRow = FRow.Row - 1
                                     tbEx.Resize tbEx.Range.Resize(tbEx.ListRows.Count + 1, tbEx.ListColumns.Count)
                                     tbEx.ListColumns(Année).DataBodyRange(EndE + 1, 1).value = tbEM.ListColumns(Année).DataBodyRange(YRow, 1).value
                                     monthdateTEST = DateValue(monthnumberD & " " & tbEM.ListColumns(Année).DataBodyRange(YRow, 1).value)
                                     tbEx.ListColumns(Mois).DataBodyRange(EndE + 1, 1).value = monthdateTEST
                                     SCodeTEST = tbEM.ListColumns(Année).DataBodyRange(YRow, 1).value & monthnumber & SE
                                     tbEx.ListColumns(Logx).DataBodyRange(EndE + 1, 1).value = SCodeTEST
                                     tbEx.ListColumns(Montant).DataBodyRange(EndE + 1, 1).value = 0
                                     EndE = EndE + 1
                                     tbTP.ListColumns("EndE").DataBodyRange(1, 1).value = EndE
                            
                                     Set Sarray = tbEx.ListColumns(Logx).DataBodyRange
                                     Else
                                     Exit Do
                                     End If
                              End If
                         End If
                         Loop
              
              Next i
     
              For i = 1 To Endi
                    monthnumber = Format(i, "00")
                    SCode = monthnumber & SE
                    
                    retryCount = 0
                    Do
                    On Error Resume Next
                    Set FRow = Sarray.Find(What:="*" & SCode, After:=tbEx.ListColumns(Logx).DataBodyRange(StartE, 1), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                        If Not FRow Is Nothing Then
                           YRow = FRow.Row - 1
                            'CalculDiff---------------------------------------------------------------------------------
                           CalculDiff
                           tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = ValAv
                           ValMS = ValAv
                           MyValMS = Replace(Format(ValAv, "0.00"), ".", Application.International(xlDecimalSeparator))
                             If PORTE_M = 1 Then
                                Set foundLabel1 = MOY.Controls(ValName(m))
                                foundLabel1.Caption = MyValMS & CRY
                                Set foundLabel2 = MOY.Controls(LabelName(m))
                                foundLabel2.ForeColor = &H80000004
                                m = m + 1
                             Else
                                For m = 0 To NBM
                                    Set foundLabel2 = MOY.Controls(LabelName(m))
                                    If Left(foundLabel2.Caption, 2) = monthnumber Then
                                       Set foundLabel1 = MOY.Controls(ValName(m))
                                       foundLabel1.Caption = MyValMS & CRY
                                       foundLabel2.ForeColor = &H80000004
                                       m = Right(foundLabel1.Name, 2)
                                       PORTE_M = 1
                                       Exit For
                                    End If
                                Next m
                            End If
                         'Check Changement----------------------------------------------------------------------------------------------------------------
                           If tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value Then
                              tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 0
                              foundLabel1.ForeColor = &H808080
                           Else
                              tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 1
                              foundLabel1.ForeColor = &H80000004
                           End If
                         
                         Exit Do
                         Else
                         'Apply Cat3 missing ----------------------------------------------------------------
                         retryCount = retryCount + 1
                              If retryCount > maxRetries Then
                              Exit Do
                              Else
                                     tbEx.ListRows.Add
                                     REF.ListColumns("Cat_1").DataBodyRange(SE, 1).Resize(1, 3).Copy
                                     tbEx.ListColumns("Cat_1").DataBodyRange(EndE + 1, 1).PasteSpecial Paste:=xlPasteValues
                                     Application.CutCopyMode = False
                                     monthnumberD = monthnumber + 0
                                     MonthTEST = Application.WorksheetFunction.Text(DateSerial(2025, monthnumberD, 1), "[$-409]mmmm")
                                     On Error Resume Next
                                     Set FRow = tbLmmmm.ListColumns("English").DataBodyRange.Find(MonthTEST, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                                     If Not FRow Is Nothing Then
                                     YRow = FRow.Row - 1
                                     tbEx.Resize tbEx.Range.Resize(tbEx.ListRows.Count + 1, tbEx.ListColumns.Count)
                                     tbEx.ListColumns(Année).DataBodyRange(EndE + 1, 1).value = tbEM.ListColumns(Année).DataBodyRange(YRow, 1).value
                                     monthdateTEST = DateValue(monthnumberD & " " & tbEM.ListColumns(Année).DataBodyRange(YRow, 1).value)
                                     tbEx.ListColumns(Mois).DataBodyRange(EndE + 1, 1).value = monthdateTEST
                                     SCodeTEST = tbEM.ListColumns(Année).DataBodyRange(YRow, 1).value & monthnumber & SE
                                     tbEx.ListColumns(Logx).DataBodyRange(EndE + 1, 1).value = SCodeTEST
                                     tbEx.ListColumns(Montant).DataBodyRange(EndE + 1, 1).value = 0
                                     EndE = EndE + 1
                                     tbTP.ListColumns("EndE").DataBodyRange(1, 1).value = EndE
                                     Set Sarray = tbEx.ListColumns(Logx).DataBodyRange
                                     Else
                                     Exit Do
                                     End If
                              End If
                         End If
                         Loop
              
              Next i

    Else
          
              'Normal following month-----------------------------------------------------------------------------------
              For i = Starti To Endi
                    monthnumber = Format(i, "00")
                    SCode = monthnumber & SE
                    retryCount = 0
                    Do
                    On Error Resume Next
                    Set FRow = Sarray.Find(What:="*" & SCode, After:=tbEx.ListColumns(Logx).DataBodyRange(StartE, 1), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                        If Not FRow Is Nothing Then
                           YRow = FRow.Row - 1
                           'CalculDiff---------------------------------------------------------------------------------
                           CalculDiff
                           tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = ValAv
                           ValMS = ValAv
                           MyValMS = Replace(Format(ValAv, "0.00"), ".", Application.International(xlDecimalSeparator))
                           If PORTE_M = 1 Then
                                Set foundLabel1 = MOY.Controls(ValName(m))
                                foundLabel1.Caption = MyValMS & CRY
                                Set foundLabel2 = MOY.Controls(LabelName(m))
                                foundLabel2.ForeColor = &H80000004
                                m = m + 1
                             Else
                                For m = 0 To NBM
                                    Set foundLabel2 = MOY.Controls(LabelName(m))
                                    If Left(foundLabel2.Caption, 2) = monthnumber Then
                                       Set foundLabel1 = MOY.Controls(ValName(m))
                                       foundLabel1.Caption = MyValMS & CRY
                                       foundLabel2.ForeColor = &H80000004
                                       m = Right(foundLabel1.Name, 2)
                                       PORTE_M = 1
                                       Exit For
                                    End If
                                Next m
                            End If
                         'Check Changement--------------------------------------------------------------------------------------------------------
                           If tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value Then
                              tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 0
                              foundLabel1.ForeColor = &H808080
                           Else
                              tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 1
                              foundLabel1.ForeColor = &H80000004
                           End If
                         
                         Exit Do
                         Else
                         'Apply Cat3 missing ----------------------------------------------------------------
                         retryCount = retryCount + 1
                              If retryCount > maxRetries Then
                              Exit Do
                              Else
                                     tbEx.ListRows.Add
                                     REF.ListColumns("Cat_1").DataBodyRange(SE, 1).Resize(1, 3).Copy
                                     tbEx.ListColumns("Cat_1").DataBodyRange(EndE + 1, 1).PasteSpecial Paste:=xlPasteValues
                                     Application.CutCopyMode = False
                                     monthnumberD = monthnumber + 0
                                     MonthTEST = Application.WorksheetFunction.Text(DateSerial(2025, monthnumberD, 1), "[$-409]mmmm")
                                     On Error Resume Next
                                     Set FRow = tbLmmmm.ListColumns("English").DataBodyRange.Find(MonthTEST, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                                     If Not FRow Is Nothing Then
                                     YRow = FRow.Row - 1
                                     tbEx.Resize tbEx.Range.Resize(tbEx.ListRows.Count + 1, tbEx.ListColumns.Count)
                                     tbEx.ListColumns(Année).DataBodyRange(EndE + 1, 1).value = tbEM.ListColumns(Année).DataBodyRange(YRow, 1).value
                                     monthdateTEST = DateValue(monthnumberD & " " & tbEM.ListColumns(Année).DataBodyRange(YRow, 1).value)
                                     tbEx.ListColumns(Mois).DataBodyRange(EndE + 1, 1).value = monthdateTEST
                                     SCodeTEST = tbEM.ListColumns(Année).DataBodyRange(YRow, 1).value & monthnumber & SE
                                     tbEx.ListColumns(Logx).DataBodyRange(EndE + 1, 1).value = SCodeTEST
                                     tbEx.ListColumns(Montant).DataBodyRange(EndE + 1, 1).value = 0
                                     EndE = EndE + 1
                                     tbTP.ListColumns("EndE").DataBodyRange(1, 1).value = EndE
                                     Set Sarray = tbEx.ListColumns(Logx).DataBodyRange
                                     Else
                                     Exit Do
                                     End If
                              End If
                         End If
                         Loop

              Next i
              
    End If

'Change average of only one month-------------------------------------------------------------------------------------------------------
ElseIf MO1 <> "" And MO2 = "" Then
B = 1

         'Reset period------------------------------------
         For i = 0 To NBM
         Set foundLabel = MOY.Controls(LabelName(i))
         foundLabel.ForeColor = &H808080
         Next i

     monthdate1 = VBA.DateValue(SpecialM1 & " " & 2000)
     Starti = Month(monthdate1)
     monthnumber = Format(Starti, "00")
     SCode = monthnumber & SE
       retryCount = 0
       Do
       On Error Resume Next
    Set FRow = Sarray.Find(What:="*" & SCode, After:=tbEx.ListColumns(Logx).DataBodyRange(StartE, 1), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
              If Not FRow Is Nothing Then
              YRow = FRow.Row - 1
              'CalculDiff---------------------------------------------------------------------------------
               CalculDiff
               tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = ValAv
               ValMS = ValAv
                MyValMS = Replace(Format(ValAv, "0.00"), ".", Application.International(xlDecimalSeparator))
                             If PORTE_M = 1 Then
                                Set foundLabel1 = MOY.Controls(ValName(m))
                                foundLabel1.Caption = MyValMS & CRY
                                Set foundLabel2 = MOY.Controls(LabelName(m))
                                foundLabel2.ForeColor = &H80000004
                                m = m + 1
                             Else
                                For m = 0 To NBM
                                    Set foundLabel2 = MOY.Controls(LabelName(m))
                                    If Left(foundLabel2.Caption, 2) = monthnumber Then
                                       Set foundLabel1 = MOY.Controls(ValName(m))
                                       foundLabel1.Caption = MyValMS & CRY
                                       foundLabel2.ForeColor = &H80000004
                                       m = Right(foundLabel1.Name, 2)
                                       PORTE_M = 1
                                       Exit For
                                    End If
                                Next m
                            End If
                           'Check changement----------------------------------------------------------------------------------------------------------------
                           If tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value Then
                              tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 0
                              foundLabel1.ForeColor = &H808080
                           Else
                              tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 1
                              foundLabel1.ForeColor = &H80000004
                           End If
                Exit Do
                Else
                         'Apply Cat3 missing ----------------------------------------------------------------
                                     retryCount = retryCount + 1
                                     If retryCount > maxRetries Then
                                     Exit Do
                                     Else
                                     tbEx.ListRows.Add
                                     REF.ListColumns("Cat_1").DataBodyRange(SE, 1).Resize(1, 3).Copy
                                     tbEx.ListColumns("Cat_1").DataBodyRange(EndE + 1, 1).PasteSpecial Paste:=xlPasteValues
                                     Application.CutCopyMode = False
                                     monthnumberD = monthnumber + 0
                                     MonthTEST = Application.WorksheetFunction.Text(DateSerial(2025, monthnumberD, 1), "[$-409]mmmm")
                                     On Error Resume Next
                                     Set FRow = tbLmmmm.ListColumns("English").DataBodyRange.Find(MonthTEST, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                                     If Not FRow Is Nothing Then
                                     YRow = FRow.Row - 1
                                     tbEx.Resize tbEx.Range.Resize(tbEx.ListRows.Count + 1, tbEx.ListColumns.Count)
                                     tbEx.ListColumns(Année).DataBodyRange(EndE + 1, 1).value = tbEM.ListColumns(Année).DataBodyRange(YRow, 1).value
                                     monthdateTEST = DateValue(monthnumberD & " " & tbEM.ListColumns(Année).DataBodyRange(YRow, 1).value)
                                     tbEx.ListColumns(Mois).DataBodyRange(EndE + 1, 1).value = monthdateTEST
                                     SCodeTEST = tbEM.ListColumns(Année).DataBodyRange(YRow, 1).value & monthnumber & SE
                                     tbEx.ListColumns(Logx).DataBodyRange(EndE + 1, 1).value = SCodeTEST
                                     tbEx.ListColumns(Montant).DataBodyRange(EndE + 1, 1).value = 0
                                     EndE = EndE + 1
                                     tbTP.ListColumns("EndE").DataBodyRange(1, 1).value = EndE
                                     Set Sarray = tbEx.ListColumns(Logx).DataBodyRange
                                     Else
                                     Exit Do
                                     End If
                                     End If
                End If
                Loop

'Enter Average starting at the begening-----------------------------------------------------------------------------------------------
ElseIf MO1 = "" And MO2 <> "" Then
B = 1
     'Reset period------------------------------------
         For i = 0 To NBM
         Set foundLabel = MOY.Controls(LabelName(i))
         foundLabel.ForeColor = &H808080
         Next i
         
     monthdate1 = tbEx.ListColumns(Mois).DataBodyRange(StartE, 1).value
     monthdate2 = VBA.DateValue(SpecialM2 & " " & 2000)
     Starti = Month(monthdate1)
     Endi = Month(monthdate2)
          If Endi < Starti Then
             Endir = MSend
          Else
             Endir = Endi
          End If
     
            C = 0
            For i = Starti To Endir
                    If C = 1 And i = Endir + 1 Then
                       Exit For
                    ElseIf C = 1 And Endi = 12 Then
                       Exit For
                    End If
                    monthnumber = Format(i, "00")
                    SCode = monthnumber & SE
                    retryCount = 0
                    Do
                    On Error Resume Next
                    Set FRow = Sarray.Find(What:="*" & SCode, After:=tbEx.ListColumns(Logx).DataBodyRange(StartE, 1), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                    On Error Resume Next
                        If Not FRow Is Nothing Then
                           YRow = FRow.Row - 1
                           'CalculDiff---------------------------------------------------------------------------------
                           CalculDiff
                           tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = ValAv
                           ValMS = ValAv
                           MyValMS = Replace(Format(ValAv, "0.00"), ".", Application.International(xlDecimalSeparator))
                            If PORTE_M = 1 Then
                                Set foundLabel1 = MOY.Controls(ValName(m))
                                foundLabel1.Caption = MyValMS & CRY
                                Set foundLabel2 = MOY.Controls(LabelName(m))
                                foundLabel2.ForeColor = &H80000004
                                m = m + 1
                             Else
                                For m = 0 To NBM
                                    Set foundLabel2 = MOY.Controls(LabelName(m))
                                    If Left(foundLabel2.Caption, 2) = monthnumber Then
                                       Set foundLabel1 = MOY.Controls(ValName(m))
                                       foundLabel1.Caption = MyValMS & CRY
                                       foundLabel2.ForeColor = &H80000004
                                       m = Right(foundLabel1.Name, 2)
                                       PORTE_M = 1
                                       Exit For
                                    End If
                                Next m
                            End If
                                    'Check Changement--------------------------------------------------------------------------------------------------------------
                                    If tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value Then
                                       tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 0
                                       foundLabel1.ForeColor = &H808080
                                    Else
                                       tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 1
                                       foundLabel1.ForeColor = &H80000004
                                    End If
                    
                    Exit Do
                Else
                         'Apply Cat3 missing ----------------------------------------------------------------
                                     retryCount = retryCount + 1
                                     If retryCount > maxRetries Then
                                     Exit Do
                                     Else
                                     tbEx.ListRows.Add
                                     REF.ListColumns("Cat_1").DataBodyRange(SE, 1).Resize(1, 3).Copy
                                     tbEx.ListColumns("Cat_1").DataBodyRange(EndE + 1, 1).PasteSpecial Paste:=xlPasteValues
                                     Application.CutCopyMode = False
                                     monthnumberD = monthnumber + 0
                                     MonthTEST = Application.WorksheetFunction.Text(DateSerial(2025, monthnumberD, 1), "[$-409]mmmm")
                                     On Error Resume Next
                                     Set FRow = tbLmmmm.ListColumns("English").DataBodyRange.Find(MonthTEST, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                                     If Not FRow Is Nothing Then
                                     YRow = FRow.Row - 1
                                     tbEx.Resize tbEx.Range.Resize(tbEx.ListRows.Count + 1, tbEx.ListColumns.Count)
                                     tbEx.ListColumns(Année).DataBodyRange(EndE + 1, 1).value = tbEM.ListColumns(Année).DataBodyRange(YRow, 1).value
                                     monthdateTEST = DateValue(monthnumberD & " " & tbEM.ListColumns(Année).DataBodyRange(YRow, 1).value)
                                     tbEx.ListColumns(Mois).DataBodyRange(EndE + 1, 1).value = monthdateTEST
                                     SCodeTEST = tbEM.ListColumns(Année).DataBodyRange(YRow, 1).value & monthnumber & SE
                                     tbEx.ListColumns(Logx).DataBodyRange(EndE + 1, 1).value = SCodeTEST
                                     tbEx.ListColumns(Montant).DataBodyRange(EndE + 1, 1).value = 0
                                     EndE = EndE + 1
                                     tbTP.ListColumns("EndE").DataBodyRange(1, 1).value = EndE
                                     Set Sarray = tbEx.ListColumns(Logx).DataBodyRange
                                     Else
                                     Exit Do
                                     End If
                                     End If
                End If
                Loop

                'Reset new year-----------------------------------------------------------------------------------------
                If i = 12 Then
                   Endir = Endi
                   C = 1
                   i = 0
                End If
             Next i

Else
Me.Average.value = ""
Exit Sub
End If

                   'Set New Value----------------------------------------------------------------------------------
                    If B = 1 Then
                    ValMS = 0
                     For i = 0 To NBM
                    Set foundLabel1 = MOY.Controls(ValName(i))
                    MyValMS = Me.foundLabel1.Caption
                    MyValMS = Replace(MyValMS, CRY, 0)
                    ValMSCount = CDbl(MyValMS)
                        If Me.foundLabel1.Caption = "" Then
                        ValMSCount = 0
                        End If
                        ValMS = ValMS + ValMSCount
                    Next i
                    TTB = ValMS
                    TTB = TTB + 0
                    tbTPC.ListColumns(Montant).DataBodyRange(SE, 1).value = TTB
                    TTB = Format(TTB, "0.00")
                    Me.TBudget.Caption = Abs(TTB) & CRY
                    TTDiff = TTB - TTREF
                    tbTPC.ListColumns("Diff").DataBodyRange(SE, 1).value = TTDiff
                    TTDiff = Format(TTDiff, "0.00")
                    MBAV = TTB / MSend
                    MBAV = Format(MBAV, "0.00")
                    Me.BAV.Caption = MBAV & CRY
                    MREF = TTREF / MSend
                    MREF = Format(MREF, "0.00")
                    Me.REFAV.Caption = MREF & CRY
                    If MREF = MBAV Then
                         Me.TBudget.ForeColor = &H808080
                         Me.BAV.ForeColor = &H808080
                       Else
                         Me.TBudget.ForeColor = &H80000004
                         Me.BAV.ForeColor = &H80000004
                      End If
                    Diff = Format(Diff, "0.00")
                    If Diff < 0 Then
                    LibR = ThisWorkbook.Sheets("Library").Range("A128").value
                    Me.FrameDiff.Caption = LibR
                    Me.LabelDiff.Caption = Diff & CRY & " " & ChrW(8600)
                    Me.LabelDiff.ForeColor = vbRed
                    Else
                    LibR = ThisWorkbook.Sheets("Library").Range("A119").value
                    Me.FrameDiff.Caption = LibR
                    Me.LabelDiff.Caption = Diff & CRY & " " & ChrW(&H2191)
                    Me.LabelDiff.ForeColor = &H80000004
                    End If
                    If TTREF = 0 Or TTREF = "" Then
                    TTP = 100
                    Else
                    TTP = ((TTB / TTREF) - 1) * 100
                    TTP = Format(TTP, "0.00")
                    End If
                    If TTDiff < 0 Then
                    LibR = ThisWorkbook.Sheets("Library").Range("A125").value
                    Me.GainTotal.Caption = LibR
                    Me.LabelTTDiff.ForeColor = &H80000004
                    Me.LabelTTDiffP.ForeColor = &H80000004
                    ElseIf TTDiff = 0 Then
                    LibR = ThisWorkbook.Sheets("Library").Range("A118").value
                    Me.GainTotal.Caption = LibR
                    Me.LabelTTDiff.ForeColor = &H808080
                    Me.LabelTTDiffP.ForeColor = &H808080
                    ElseIf TTDiff > 0 Then
                    LibR = ThisWorkbook.Sheets("Library").Range("A118").value
                    Me.GainTotal.Caption = LibR
                    Me.LabelTTDiff.ForeColor = &H80000004
                    Me.LabelTTDiffP.ForeColor = &H80000004
                    End If
                    Me.LabelTTDiff.Caption = TTDiff & CRY
                    If TTP < 0 Then
                    Me.LabelTTDiffP.Caption = "(" & TTP & "%" & ")"
                    Else
                    Me.LabelTTDiffP.Caption = "(+" & TTP & "%" & ")"
                    End If
                    Me.Average.value = ""
                    Me.CashFlow.Caption = Format(tbTP.ListColumns("StartE").DataBodyRange(2, 1).value, "0.00") & CRY
                    End If
                    

End Sub

Sub CalculDiff()
Dim AverageL As Double
NewVal = 0


If P = 1 Then
   AverageL = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value
   'Choice Pondlar-----------------------------------------------------------------------
   MPond = tbEx.ListColumns("Ratio").DataBodyRange(YRow, 1).value
   NewVal = MPond * ValAv

ElseIf P = 2 Then
   'Choice Pond-----------------------------------------------------------------------
    AverageL = tbEx.ListColumns("Montant").DataBodyRange(YRow, 1).value

        MPond = 1 + ValP2

     NewVal = AverageL * MPond
      
ElseIf P = 3 Then
 NewVal = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value

 AverageL = tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value

ElseIf P = 0 Then
'Choice Average--------------------------------------------------------------------------
 AverageL = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value
 NewVal = ValAv
End If
                                                      
 
  If AverageL = NewVal Then
     Diff = Diff + 0
  Else
     Diff = NewVal - AverageL + Diff
  End If

End Sub

Private Sub CommandP_Click()
Dim LibR As String

m = 0
PORTE_M = 0
P = 0
Diff = 0
F = 0
MO1 = Me.MO_1.value

                    Set FRow = tbMonth.ListColumns("Mois").DataBodyRange.Find(What:=MO1, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                    On Error Resume Next
                        If Not FRow Is Nothing Then
                           SpecialM1 = FRow.Row - 1
                           MO1 = tbLmmmm.ListColumns("English").DataBodyRange(SpecialM1, 1).value
                        End If
                  
MO2 = Me.MO_2.value

                   Set FRow = tbMonth.ListColumns("Mois").DataBodyRange.Find(What:=MO2, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                    On Error Resume Next
                        If Not FRow Is Nothing Then
                           SpecialM2 = FRow.Row - 1
                           MO2 = tbLmmmm.ListColumns("English").DataBodyRange(SpecialM2, 1).value
                        End If
                                       
    ValP1 = Me.Average.value
    ValP1 = Replace(ValP1, localDec, altDec)
    If vbaSep = "." Then
    ValP1 = Replace(ValP1, altDec, Application.International(xlDecimalSeparator))
    End If
    
    If IsNumeric(ValP1) Then
        ValP1 = ValP1 + 0
    Else
        ValP1 = ""
        If Me.Average.value = "" Then
        Me.Average.BorderStyle = fmBorderStyleSingle
        Me.Average.BorderColor = vbRed
        Else
        Me.Average.ForeColor = vbRed
        End If
        ValAv = ""
    End If
    
    ValP2 = Me.Pond.value
    ValP2 = Replace(ValP2, localDec, altDec)
    If vbaSep = "." Then
    ValP2 = Replace(ValP2, altDec, Application.International(xlDecimalSeparator))
    End If

    If IsNumeric(ValP2) Then
       ValP2 = ValP2 + 0
       ValP2 = ValP2 / 100
    Else
        ValP2 = ""
        If Me.Pond.value = "" Then
        Me.Pond.BorderStyle = fmBorderStyleSingle
        Me.Pond.BorderColor = vbRed
        Else
        Me.Pond.ForeColor = vbRed
        End If
        ValAv = ""
    End If
    
    If ValP1 = "" And ValP2 = "" Then
       F = 2
       ValAv = ""
    ElseIf ValP1 <> "" And ValP2 = "" Then
       P = 1
       ValAv = ValP1
    ElseIf ValP1 = "" And ValP2 <> "" Then
       P = 2
       ValAv = ValP2
    End If
    If Me.Average.value <> "" And Me.Pond.value <> "" Then
       Me.Average.ForeColor = vbRed
       Me.Pond.ForeColor = vbRed
       Exit Sub
    ElseIf Me.Average.value <> "" And Me.Pond.value = "" Then
       Me.Pond.BorderStyle = fmBorderStyleNone
       Me.Pond.ForeColor = vbRed
       F = 2
    ElseIf Me.Average.value = "" And Me.Pond.value <> "" Then
       Me.Average.BorderStyle = fmBorderStyleNone
       Me.Average.ForeColor = vbRed
       F = 2
    End If
       
    
Set Sarray = tbEx.ListColumns(Logx).DataBodyRange
If Me.Cat = "" Then
Me.Cat.BorderStyle = fmBorderStyleSingle
Me.Cat.BorderColor = vbRed
   F = 2
End If
If MO1 = "" And MO2 = "" Then
Me.MO_1.BorderStyle = fmBorderStyleSingle
Me.MO_1.BorderColor = vbRed
Me.MO_2.BorderStyle = fmBorderStyleSingle
Me.MO_2.BorderColor = vbRed
   F = 2
End If
  

If (MO1 = "" And MO2 = "") Or Me.Cat = "" Or ValAv = "" Then
B = 0
    If F = 2 Then
       Exit Sub
    End If

'Enter ponderation value in an Array--------------------------------------------------------------------------------------------
ElseIf MO1 <> "" And MO2 <> "" Then
B = 1

        'Reset period------------------------------------
         For i = 0 To NBM
         Set foundLabel = MOY.Controls(LabelName(i))
         foundLabel.ForeColor = &H808080
         Next i

     monthdate1 = VBA.DateValue(SpecialM1 & " " & 2000)
     monthdate2 = VBA.DateValue(SpecialM2 & " " & 2000)
     Starti = Month(monthdate1)
     Endi = Month(monthdate2)
     
     If Endi < Starti Then
     
            'Changing year--------------------------------------------------------------------------------------------
              For i = Starti To MSend
                    monthnumber = Format(i, "00")
                    SCode = monthnumber & SE
                    
                    Set FRow = Sarray.Find(What:="*" & SCode, After:=tbEx.ListColumns(Logx).DataBodyRange(StartE, 1), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                    On Error Resume Next
                        If Not FRow Is Nothing Then
                           YRow = FRow.Row - 1
                               If ValP1 <> "" Then
                                 'CalculDiff---------------------------------------------------------------------------------
                                  CalculDiff
                                  Valup = Replace(NewVal, ",", ".")
                                  tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = Valup
                                  ValMS = CDbl(NewVal)
                                  MyValMS = Replace(Format(ValMS, "0.00"), ".", Application.International(xlDecimalSeparator))
                                  If PORTE_M = 1 Then
                                     Set foundLabel1 = MOY.Controls(ValName(m))
                                     foundLabel1.Caption = MyValMS & CRY
                                     Set foundLabel2 = MOY.Controls(LabelName(m))
                                     foundLabel2.ForeColor = &H80000004
                                     m = m + 1
                                  Else
                                     For m = 0 To NBM
                                       Set foundLabel2 = MOY.Controls(LabelName(m))
                                       If Left(foundLabel2.Caption, 2) = monthnumber Then
                                       Set foundLabel1 = MOY.Controls(ValName(m))
                                       foundLabel1.Caption = MyValMS & CRY
                                       foundLabel2.ForeColor = &H80000004
                                       m = Right(foundLabel1.Name, 2)
                                       PORTE_M = 1
                                       Exit For
                                       End If
                                    Next m
                                 End If
                               Else
                                  'CalculDiff---------------------------------------------------------------------------------
                                  CalculDiff
                                  Valup = Replace(NewVal, ",", ".")
                                  tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = Valup
                                  ValMS = CDbl(NewVal)
                                  MyValMS = Replace(Format(ValMS, "0.00"), ".", Application.International(xlDecimalSeparator))
                                  If PORTE_M = 1 Then
                                     Set foundLabel1 = MOY.Controls(ValName(m))
                                     foundLabel1.Caption = MyValMS & CRY
                                     Set foundLabel2 = MOY.Controls(LabelName(m))
                                     foundLabel2.ForeColor = &H80000004
                                     m = m + 1
                                  Else
                                     For m = 0 To NBM
                                       Set foundLabel2 = MOY.Controls(LabelName(m))
                                       If Left(foundLabel2.Caption, 2) = monthnumber Then
                                       Set foundLabel1 = MOY.Controls(ValName(m))
                                       foundLabel1.Caption = MyValMS & CRY
                                       foundLabel2.ForeColor = &H80000004
                                       m = Right(foundLabel1.Name, 2)
                                       PORTE_M = 1
                                       Exit For
                                       End If
                                    Next m
                                 End If
                               End If
                                
                         End If
                         'Check Changement--------------------------------------------------------------------------------------------------------------
                                    If tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value Then
                                       tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 0
                                       foundLabel1.ForeColor = &H808080
                                    Else
                                       tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 1
                                       foundLabel1.ForeColor = &H80000004
                                    End If
              
              Next i
     
              For i = 1 To Endi
                    monthnumber = Format(i, "00")
                    SCode = monthnumber & SE
                    
                    Set FRow = Sarray.Find(What:="*" & SCode, After:=tbEx.ListColumns(Logx).DataBodyRange(StartE, 1), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                    On Error Resume Next
                        If Not FRow Is Nothing Then
                           YRow = FRow.Row - 1
                               If ValP1 <> "" Then
                                 'CalculDiff---------------------------------------------------------------------------------
                                  CalculDiff
                                  Valup = Replace(NewVal, ",", ".")
                                  tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = Valup
                                  ValMS = CDbl(NewVal)
                                  MyValMS = Replace(Format(ValMS, "0.00"), ".", Application.International(xlDecimalSeparator))
                                  If PORTE_M = 1 Then
                                     Set foundLabel1 = MOY.Controls(ValName(m))
                                     foundLabel1.Caption = MyValMS & CRY
                                     Set foundLabel2 = MOY.Controls(LabelName(m))
                                     foundLabel2.ForeColor = &H80000004
                                     m = m + 1
                                  Else
                                     For m = 0 To NBM
                                       Set foundLabel2 = MOY.Controls(LabelName(m))
                                       If Left(foundLabel2.Caption, 2) = monthnumber Then
                                       Set foundLabel1 = MOY.Controls(ValName(m))
                                       foundLabel1.Caption = MyValMS & CRY
                                       foundLabel2.ForeColor = &H80000004
                                       m = Right(foundLabel1.Name, 2)
                                       PORTE_M = 1
                                       Exit For
                                       End If
                                    Next m
                                 End If
                               Else
                                 'CalculDiff---------------------------------------------------------------------------------
                                  CalculDiff
                                  Valup = Replace(NewVal, ",", ".")
                                  tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = Valup
                                  ValMS = CDbl(NewVal)
                                  MyValMS = Replace(Format(ValMS, "0.00"), ".", Application.International(xlDecimalSeparator))
                                  If PORTE_M = 1 Then
                                     Set foundLabel1 = MOY.Controls(ValName(m))
                                     foundLabel1.Caption = MyValMS & CRY
                                     Set foundLabel2 = MOY.Controls(LabelName(m))
                                     foundLabel2.ForeColor = &H80000004
                                     m = m + 1
                                  Else
                                     For m = 0 To NBM
                                       Set foundLabel2 = MOY.Controls(LabelName(m))
                                       If Left(foundLabel2.Caption, 2) = monthnumber Then
                                       Set foundLabel1 = MOY.Controls(ValName(m))
                                       foundLabel1.Caption = MyValMS & CRY
                                       foundLabel2.ForeColor = &H80000004
                                       m = Right(foundLabel1.Name, 2)
                                       PORTE_M = 1
                                       Exit For
                                       End If
                                    Next m
                                 End If
                               End If
                                
                         End If
              
                        'Check Changement--------------------------------------------------------------------------------------------------------------
                                    If tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value Then
                                       tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 0
                                       foundLabel1.ForeColor = &H808080
                                    Else
                                       tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 1
                                       foundLabel1.ForeColor = &H80000004
                                    End If
              
              Next i

     Else
          
              'Normal following month-----------------------------------------------------------------------------------
              For i = Starti To Endi
                    monthnumber = Format(i, "00")
                    SCode = monthnumber & SE
                    
                    Set FRow = Sarray.Find(What:="*" & SCode, After:=tbEx.ListColumns(Logx).DataBodyRange(StartE, 1), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                    On Error Resume Next
                        If Not FRow Is Nothing Then
                           YRow = FRow.Row - 1
                               If ValP1 <> "" Then
                                 'CalculDiff---------------------------------------------------------------------------------
                                  CalculDiff
                                  Valup = Replace(NewVal, ",", ".")
                                  tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = Valup
                                  ValMS = CDbl(NewVal)
                                  MyValMS = Replace(Format(ValMS, "0.00"), ".", Application.International(xlDecimalSeparator))
                                  If PORTE_M = 1 Then
                                     Set foundLabel1 = MOY.Controls(ValName(m))
                                     foundLabel1.Caption = MyValMS & CRY
                                     Set foundLabel2 = MOY.Controls(LabelName(m))
                                     foundLabel2.ForeColor = &H80000004
                                     m = m + 1
                                  Else
                                     For m = 0 To NBM
                                       Set foundLabel2 = MOY.Controls(LabelName(m))
                                       If Left(foundLabel2.Caption, 2) = monthnumber Then
                                       Set foundLabel1 = MOY.Controls(ValName(m))
                                       foundLabel1.Caption = MyValMS & CRY
                                       foundLabel2.ForeColor = &H80000004
                                       m = Right(foundLabel1.Name, 2)
                                       PORTE_M = 1
                                       Exit For
                                       End If
                                    Next m
                                 End If
                               Else
                                 'CalculDiff---------------------------------------------------------------------------------
                                  CalculDiff
                                  Valup = Replace(NewVal, ",", ".")
                                  tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = Valup
                                  ValMS = CDbl(NewVal)
                                  MyValMS = Replace(Format(ValMS, "0.00"), ".", Application.International(xlDecimalSeparator))
                                  If PORTE_M = 1 Then
                                     Set foundLabel1 = MOY.Controls(ValName(m))
                                     foundLabel1.Caption = MyValMS & CRY
                                     Set foundLabel2 = MOY.Controls(LabelName(m))
                                     foundLabel2.ForeColor = &H80000004
                                     m = m + 1
                                  Else
                                     For m = 0 To NBM
                                       Set foundLabel2 = MOY.Controls(LabelName(m))
                                       If Left(foundLabel2.Caption, 2) = monthnumber Then
                                       Set foundLabel1 = MOY.Controls(ValName(m))
                                       foundLabel1.Caption = MyValMS & CRY
                                       foundLabel2.ForeColor = &H80000004
                                       m = Right(foundLabel1.Name, 2)
                                       PORTE_M = 1
                                       Exit For
                                       End If
                                    Next m
                                 End If
                               End If
                                
                         End If
                         'Check Changement--------------------------------------------------------------------------------------------------------------
                                    If tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value Then
                                       tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 0
                                       foundLabel1.ForeColor = &H808080
                                    Else
                                       tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 1
                                       foundLabel1.ForeColor = &H80000004
                                    End If
              
              Next i

    End If

'Change ponderation value of only one month-------------------------------------------------------------------------------------------------------
ElseIf MO1 <> "" And MO2 = "" Then
B = 1

         'Reset period------------------------------------
         For i = 0 To NBM
         Set foundLabel = MOY.Controls(LabelName(i))
         foundLabel.ForeColor = &H808080
         Next i

     monthdate1 = VBA.DateValue(SpecialM1 & " " & 2000)
     Starti = Month(monthdate1)
     monthnumber = Format(Starti, "00")
     SCode = monthnumber & SE
    Set FRow = Sarray.Find(What:="*" & SCode, After:=tbEx.ListColumns(Logx).DataBodyRange(StartE, 1), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        On Error Resume Next
           If Not FRow Is Nothing Then
              YRow = FRow.Row - 1
                               If ValP1 <> "" Then
                                 'CalculDiff---------------------------------------------------------------------------------
                                  CalculDiff
                                  Valup = Replace(NewVal, ",", ".")
                                  tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = Valup
                                  ValMS = CDbl(NewVal)
                                  MyValMS = Replace(Format(ValMS, "0.00"), ".", Application.International(xlDecimalSeparator))
                                  If PORTE_M = 1 Then
                                     Set foundLabel1 = MOY.Controls(ValName(m))
                                     foundLabel1.Caption = MyValMS & CRY
                                     Set foundLabel2 = MOY.Controls(LabelName(m))
                                     foundLabel2.ForeColor = &H80000004
                                     m = m + 1
                                  Else
                                     For m = 0 To NBM
                                       Set foundLabel2 = MOY.Controls(LabelName(m))
                                       If Left(foundLabel2.Caption, 2) = monthnumber Then
                                       Set foundLabel1 = MOY.Controls(ValName(m))
                                       foundLabel1.Caption = MyValMS & CRY
                                       foundLabel2.ForeColor = &H80000004
                                       m = Right(foundLabel1.Name, 2)
                                       PORTE_M = 1
                                       Exit For
                                       End If
                                    Next m
                                 End If
                               Else
                                 'CalculDiff---------------------------------------------------------------------------------
                                  CalculDiff
                                  Valup = Replace(NewVal, ",", ".")
                                  tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = Valup
                                  ValMS = CDbl(NewVal)
                                  MyValMS = Replace(Format(ValMS, "0.00"), ".", Application.International(xlDecimalSeparator))
                                  If PORTE_M = 1 Then
                                     Set foundLabel1 = MOY.Controls(ValName(m))
                                     foundLabel1.Caption = MyValMS & CRY
                                     Set foundLabel2 = MOY.Controls(LabelName(m))
                                     foundLabel2.ForeColor = &H80000004
                                     m = m + 1
                                  Else
                                     For m = 0 To NBM
                                       Set foundLabel2 = MOY.Controls(LabelName(m))
                                       If Left(foundLabel2.Caption, 2) = monthnumber Then
                                       Set foundLabel1 = MOY.Controls(ValName(m))
                                       foundLabel1.Caption = MyValMS & CRY
                                       foundLabel2.ForeColor = &H80000004
                                       m = Right(foundLabel1.Name, 2)
                                       PORTE_M = 1
                                       Exit For
                                       End If
                                    Next m
                                 End If
                               End If
                        'Check Changement--------------------------------------------------------------------------------------------------------------
                                    If tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value Then
                                       tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 0
                                       foundLabel1.ForeColor = &H808080
                                    Else
                                       tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 1
                                       foundLabel1.ForeColor = &H80000004
                                    End If
            End If

'Enter ponderation value  starting at the begening-----------------------------------------------------------------------------------------------
ElseIf MO1 = "" And MO2 <> "" Then
B = 1

        'Reset period------------------------------------
         For i = 0 To NBM
         Set foundLabel = MOY.Controls(LabelName(i))
         foundLabel.ForeColor = &H808080
         Next i

     monthdate1 = tbEx.ListColumns(Mois).DataBodyRange(StartE, 1).value
     monthdate2 = VBA.DateValue(SpecialM2 & " " & 2000)
     Starti = Month(monthdate1)
     Endi = Month(monthdate2)
          If Endi < Starti Then
             Endir = 12
          Else
             Endir = Endi
          End If
     
            C = 0
            For i = Starti To Endir
                    If C = 1 And i = Endir + 1 Then
                       Exit For
                    ElseIf C = 1 And Endi = 12 Then
                       Exit For
                    End If
                    monthnumber = Format(i, "00")
                    SCode = monthnumber & SE
                    
                    Set FRow = Sarray.Find(What:="*" & SCode, After:=tbEx.ListColumns(Logx).DataBodyRange(StartE, 1), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                    On Error Resume Next
                        If Not FRow Is Nothing Then
                           YRow = FRow.Row - 1
                               If ValP1 <> "" Then
                                  'CalculDiff---------------------------------------------------------------------------------
                                  CalculDiff
                                  Valup = Replace(NewVal, ",", ".")
                                  tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = Valup
                                  ValMS = CDbl(NewVal)
                                  MyValMS = Replace(Format(ValMS, "0.00"), ".", Application.International(xlDecimalSeparator))
                                  If PORTE_M = 1 Then
                                     Set foundLabel1 = MOY.Controls(ValName(m))
                                     foundLabel1.Caption = MyValMS & CRY
                                     Set foundLabel2 = MOY.Controls(LabelName(m))
                                     foundLabel2.ForeColor = &H80000004
                                     m = m + 1
                                  Else
                                     For m = 0 To NBM
                                       Set foundLabel2 = MOY.Controls(LabelName(m))
                                       If Left(foundLabel2.Caption, 2) = monthnumber Then
                                       Set foundLabel1 = MOY.Controls(ValName(m))
                                       foundLabel1.Caption = MyValMS & CRY
                                       foundLabel2.ForeColor = &H80000004
                                       m = Right(foundLabel1.Name, 2)
                                       PORTE_M = 1
                                       Exit For
                                       End If
                                    Next m
                                 End If
                               Else
                                 'CalculDiff---------------------------------------------------------------------------------
                                  CalculDiff
                                  Valup = Replace(NewVal, ",", ".")
                                  tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = Valup
                                  ValMS = CDbl(NewVal)
                                  MyValMS = Replace(Format(ValMS, "0.00"), ".", Application.International(xlDecimalSeparator))
                                  If PORTE_M = 1 Then
                                     Set foundLabel1 = MOY.Controls(ValName(m))
                                     foundLabel1.Caption = MyValMS & CRY
                                     Set foundLabel2 = MOY.Controls(LabelName(m))
                                     foundLabel2.ForeColor = &H80000004
                                     m = m + 1
                                  Else
                                     For m = 0 To NBM
                                       Set foundLabel2 = MOY.Controls(LabelName(m))
                                       If Left(foundLabel2.Caption, 2) = monthnumber Then
                                       Set foundLabel1 = MOY.Controls(ValName(m))
                                       foundLabel1.Caption = MyValMS & CRY
                                       foundLabel2.ForeColor = &H80000004
                                       m = Right(foundLabel1.Name, 2)
                                       PORTE_M = 1
                                       Exit For
                                       End If
                                    Next m
                                 End If
                               End If
                                
                         End If
                'Reset new year-----------------------------------------------------------------------------------------
                If i = 12 Then
                   Endir = Endi
                   C = 1
                   i = 0
                End If
                 'Check Changement--------------------------------------------------------------------------------------------------------------
                                    If tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value Then
                                       tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 0
                                       foundLabel1.ForeColor = &H808080
                                    Else
                                       tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 1
                                       foundLabel1.ForeColor = &H80000004
                                    End If
             
             Next i
             
Else
P = 0
Me.Average.value = ""
Me.Pond.value = ""
Exit Sub
End If

                    If B = 1 Then
                    ValMS = 0
                     For i = 0 To NBM
                    Set foundLabel1 = MOY.Controls(ValName(i))
                    MyValMS = Me.foundLabel1.Caption
                    MyValMS = Replace(MyValMS, CRY, 0)
                    ValMSCount = CDbl(MyValMS)
                        If Me.foundLabel1.Caption = "" Then
                        ValMSCount = 0
                        End If
                        ValMS = ValMS + ValMSCount
                    Next i
                    TTB = ValMS
                    tbTPC.ListColumns(Montant).DataBodyRange(SE, 1).value = TTB
                    TTB = Format(TTB, "0.00")
                    MyTTB = Abs(TTB) & CRY
                    MyTTB = Replace(MyTTB, altDec, localDec)
                    Me.TBudget.Caption = MyTTB
                    TTDiff = TTB - TTREF
                    tbTPC.ListColumns("Diff").DataBodyRange(SE, 1).value = TTDiff
                    MBAV = TTB / MSend
                    MBAV = Format(MBAV, "0.00")
                    MyMBAV = Replace(MBAV, altDec, localDec)
                    Me.BAV.Caption = MyMBAV & CRY
                    MREF = TTREF / MSend
                    MREF = Format(MREF, "0.00")
                    Me.REFAV.Caption = MREF & CRY
                    If MREF = MBAV Then
                         Me.TBudget.ForeColor = &H808080
                         Me.BAV.ForeColor = &H808080
                       Else
                         Me.TBudget.ForeColor = &H80000004
                         Me.BAV.ForeColor = &H80000004
                      End If
                    If TTREF = 0 Or TTREF = "" Then
                    TTP = 100
                    Else
                    TTP = ((TTB / TTREF) - 1) * 100
                    TTP = Format(TTP, "0.00")
                    End If
                    If TTDiff < 0 Then
                    LibR = ThisWorkbook.Sheets("Library").Range("A125").value
                    Me.GainTotal.Caption = LibR
                    Me.LabelTTDiff.ForeColor = &H80000004
                    Me.LabelTTDiffP.ForeColor = &H80000004
                    ElseIf TTDiff = 0 Then
                    LibR = ThisWorkbook.Sheets("Library").Range("A118").value
                    Me.GainTotal.Caption = LibR
                    Me.LabelTTDiff.ForeColor = &H808080
                    Me.LabelTTDiffP.ForeColor = &H808080
                    ElseIf TTDiff > 0 Then
                    LibR = ThisWorkbook.Sheets("Library").Range("A118").value
                    Me.GainTotal.Caption = LibR
                    Me.LabelTTDiff.ForeColor = &H80000004
                    Me.LabelTTDiffP.ForeColor = &H80000004
                    End If
                    MyTTDiff = Format(TTDiff, "0.00")
                    MyTTDiff = Replace(MyTTDiff, altDec, localDec)
                    Me.LabelTTDiff.Caption = MyTTDiff & CRY
                    If TTP < 0 Then
                    Me.LabelTTDiffP.Caption = "(" & TTP & "%" & ")"
                    Else
                    Me.LabelTTDiffP.Caption = "(+" & TTP & "%" & ")"
                    End If
                    Diff = Format(Diff, "0.00")
                    If Diff < 0 Then
                    LibR = ThisWorkbook.Sheets("Library").Range("A128").value
                    Me.FrameDiff.Caption = LibR
                    Me.LabelDiff.Caption = Diff & CRY & " " & ChrW(8600)
                    Me.LabelDiff.ForeColor = vbRed
                    Else
                    LibR = ThisWorkbook.Sheets("Library").Range("A119").value
                    Me.FrameDiff.Caption = LibR
                    Me.LabelDiff.Caption = Diff & CRY & " " & ChrW(&H2191)
                    Me.LabelDiff.ForeColor = &H80000004
                    End If
                    End If
                    Me.Average.value = ""
                    Me.Pond.value = ""
                    Me.CashFlow.Caption = Format(tbTP.ListColumns("StartE").DataBodyRange(2, 1).value, "0.00") & CRY
 
P = 0
End Sub

Private Sub Commandquit_Click()

Call Build_Budget.Budgetisation(BDG)

Unload Me
End Sub

Private Sub Fermer_Click()
Unload Me
End Sub

Private Sub Frame8_Click()

End Sub

Private Sub Label57_Click()

End Sub

Private Sub MO_1_Change()
Me.MO_1.BorderStyle = fmBorderStyleNone
Me.MO_2.BorderStyle = fmBorderStyleNone
MS = Me.MO_1.value
If MS <> "" Then

   If MS = Me.MO_2.value Then
      Me.MO_2.value = ""
      Exit Sub
   End If
   
Me.MO_2.RowSource = ""
Me.MO_2.Clear
Me.MO_2.value = ""

Set FRow = tbEM.ListColumns(Mois).DataBodyRange.Find(What:=MS, LookIn:=xlValues, LookAt:=xlWhole)
       On Error Resume Next
          If Not FRow Is Nothing Then
          YRow = FRow.Row - 1
          EMend = tbEM.ListRows.Count
          
               If YRow = EMend - 1 Then
                  Me.MO_2.value = ""
                  Exit Sub
               End If
          
          End If

Set rng = tbEM.ListColumns(Mois).DataBodyRange.Offset(YRow, 0).Resize(EMend)
For Each cell In rng
    If cell.value = "" Then
       Exit For
    End If
MOY.MO_2.AddItem cell.value

Next cell

Else
Me.MO_2.RowSource = ""
Me.MO_2.Clear
For Each cell In tbEM.ListColumns(Mois).DataBodyRange
    
    If cell.value = "" Then
       Exit For
    End If
MOY.MO_2.AddItem cell.value

Next cell

End If
End Sub

Private Sub MO_2_Change()
Me.MO_2.BorderStyle = fmBorderStyleNone
Me.MO_1.BorderStyle = fmBorderStyleNone
End Sub
Private Sub Pond_Change()
If DisableEvents Then
Exit Sub
End If
Me.Pond.ForeColor = vbWhite
Me.Pond.BorderStyle = fmBorderStyleNone
DisableEvents = True
Me.Average.value = ""
DisableEvents = False
Me.Average.BorderStyle = fmBorderStyleNone
Me.Average.ForeColor = vbWhite
End Sub
Private Sub ResetAll_Click()
Dim result As VbMsgBoxResult
Dim LibR, LibR1, LibR2 As String
LibR1 = ThisWorkbook.Sheets("Library").Range("A126").value
LibR2 = ThisWorkbook.Sheets("Library").Range("A127").value
result = MsgBox(LibR1, vbYesNo + vbExclamation, LibR2)

If result = vbYes Then

tbTPC.ListColumns("Diff").DataBodyRange.Clear
tbEx.ListColumns("Check").DataBodyRange.Clear
tbEx.ListColumns(MontantE).DataBodyRange = tbEx.ListColumns(Montant).DataBodyRange.value
tbTP.ListColumns(Montant).DataBodyRange.Copy
tbTPC.ListColumns(Montant).DataBodyRange(1, 1).PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False
tbTPC.ListColumns(Montant).DataBodyRange.NumberFormat = "0.00" & CRY

For i = 0 To 11
    Set foundLabel = MOY.Controls(ValName(i))
    foundLabel.ForeColor = &H808080
    foundLabel.Caption = ""
    Set foundLabel = MOY.Controls(ValRName(i))
    foundLabel.Caption = ""
    Set foundLabel = MOY.Controls(LabelName(i))
    foundLabel.ForeColor = &H808080
Next i


Me.MO_1.value = ""
Me.MO_2.value = ""
Me.Average.value = ""
Me.Pond.value = ""
Me.LabelDiff.Caption = ""
Me.LabelTTDiffP.Caption = ""
Me.LabelTTDiff.Caption = ""
Me.TBudget.Caption = ""
Me.TREF.Caption = ""
Me.BAV.Caption = ""
Me.REFAV.Caption = ""
Me.Average.value = ""
LibR = ThisWorkbook.Sheets("Library").Range("A119").value
Me.FrameDiff.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A118").value
Me.GainTotal.Caption = LibR
Me.CashFlow.Caption = Format(tbTP.ListColumns("StartE").DataBodyRange(2, 1).value, "0.00") & CRY

Application.EnableEvents = False
Me.Cat.value = ""
Application.EnableEvents = True

Me.Average.ForeColor = vbWhite
Me.Pond.ForeColor = vbWhite
Me.Average.BorderStyle = fmBorderStyleNone
Me.Pond.BorderStyle = fmBorderStyleNone
Me.Cat.BorderStyle = fmBorderStyleNone
Me.MO_1.BorderStyle = fmBorderStyleNone
Me.MO_2.BorderStyle = fmBorderStyleNone


End If
End Sub

Private Sub ResetM_Click()
Dim LibR As String

MO1 = Me.MO_1.value

                    Set FRow = tbMonth.ListColumns("Mois").DataBodyRange.Find(What:=MO1, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                    On Error Resume Next
                        If Not FRow Is Nothing Then
                           SpecialM1 = FRow.Row - 1
                           MO1 = tbLmmmm.ListColumns("English").DataBodyRange(SpecialM1, 1).value
                        End If
                  
MO2 = Me.MO_2.value

                   Set FRow = tbMonth.ListColumns("Mois").DataBodyRange.Find(What:=MO2, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                    On Error Resume Next
                        If Not FRow Is Nothing Then
                           SpecialM2 = FRow.Row - 1
                           MO2 = tbLmmmm.ListColumns("English").DataBodyRange(SpecialM2, 1).value
                        End If
                    
m = 0
PORTE_M = 0
P = 0
Diff = 0
Set Sarray = tbEx.ListColumns(Logx).DataBodyRange

If (MO1 = "" And MO2 = "") Or Me.Cat = "" Then
B = 0
       If MO1 = "" And MO2 = "" Then
       Me.MO_1.BorderStyle = fmBorderStyleSingle
       Me.MO_1.BorderColor = vbRed
       Me.MO_2.BorderStyle = fmBorderStyleSingle
       Me.MO_2.BorderColor = vbRed
       F = 2
       End If
       If Me.Cat = "" Then
       Me.Cat.BorderStyle = fmBorderStyleSingle
       Me.Cat.BorderColor = vbRed
       F = 2
       End If
       If F = 2 Then
       Exit Sub
       End If
'Enter Reset in an Array--------------------------------------------------------------------------------------------
ElseIf MO1 <> "" And MO2 <> "" Then
P = 3
B = 1

        'Reset period------------------------------------
         For i = 0 To NBM
         Set foundLabel = MOY.Controls(LabelName(i))
         foundLabel.ForeColor = &H808080
         Next i

     monthdate1 = VBA.DateValue(SpecialM1 & " " & 2000)
     monthdate2 = VBA.DateValue(SpecialM2 & " " & 2000)
     Starti = Month(monthdate1)
     Endi = Month(monthdate2)
     
     If Endi < Starti Then
     
            'Changing year--------------------------------------------------------------------------------------------
              For i = Starti To MSend
                    monthnumber = Format(i, "00")
                    SCode = monthnumber & SE
                    
                    Set FRow = Sarray.Find(What:="*" & SCode, After:=tbEx.ListColumns(Logx).DataBodyRange(StartE, 1), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                    On Error Resume Next
                        If Not FRow Is Nothing Then
                           YRow = FRow.Row - 1
                           'CalculDiff---------------------------------------------------------------------------------
                            CalculDiff
                           tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value
                           ValMS = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value
                           MyValMS = Format(ValMS, "0.00")
                           MyValMS = Replace(MyValMS, altDec, localDec)
                          If PORTE_M = 1 Then
                                     Set foundLabel1 = MOY.Controls(ValName(m))
                                     foundLabel1.Caption = MyValMS & CRY
                                     Set foundLabel2 = MOY.Controls(LabelName(m))
                                     foundLabel2.ForeColor = &H80000004
                                     m = m + 1
                                  Else
                                     For m = 0 To NBM
                                       Set foundLabel2 = MOY.Controls(LabelName(m))
                                       If Left(foundLabel2.Caption, 2) = monthnumber Then
                                       Set foundLabel1 = MOY.Controls(ValName(m))
                                       foundLabel1.Caption = MyValMS & CRY
                                       foundLabel2.ForeColor = &H80000004
                                       m = Right(foundLabel1.Name, 2)
                                       PORTE_M = 1
                                       Exit For
                                       End If
                                    Next m
                                 End If
                         End If
                  
              tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 0
              foundLabel1.ForeColor = &H808080
              Next i
     
              For i = 1 To Endi
                    monthnumber = Format(i, "00")
                    SCode = monthnumber & SE
                    
                    Set FRow = Sarray.Find(What:="*" & SCode, After:=tbEx.ListColumns(Logx).DataBodyRange(StartE, 1), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                    On Error Resume Next
                        If Not FRow Is Nothing Then
                           YRow = FRow.Row - 1
                           'CalculDiff---------------------------------------------------------------------------------
                            CalculDiff
                           tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value
                           ValMS = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value
                           MyValMS = Format(ValMS, "0.00")
                           MyValMS = Replace(MyValMS, altDec, localDec)
                          If PORTE_M = 1 Then
                                     Set foundLabel1 = MOY.Controls(ValName(m))
                                     foundLabel1.Caption = MyValMS & CRY
                                     Set foundLabel2 = MOY.Controls(LabelName(m))
                                     foundLabel2.ForeColor = &H80000004
                                     m = m + 1
                                  Else
                                     For m = 0 To NBM
                                       Set foundLabel2 = MOY.Controls(LabelName(m))
                                       If Left(foundLabel2.Caption, 2) = monthnumber Then
                                       Set foundLabel1 = MOY.Controls(ValName(m))
                                       foundLabel1.Caption = MyValMS & CRY
                                       foundLabel2.ForeColor = &H80000004
                                       m = Right(foundLabel1.Name, 2)
                                       PORTE_M = 1
                                       Exit For
                                       End If
                                    Next m
                                 End If
                         End If
              
                         'Check Changement--------------------------------------------------------------------------------------------------------------
                                    
                                       tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 0
                                       foundLabel1.ForeColor = &H808080
                                   
              
              Next i

     Else
          
              'Normal following month-----------------------------------------------------------------------------------
              For i = Starti To Endi
                    monthnumber = Format(i, "00")
                    SCode = monthnumber & SE
                    
                    Set FRow = Sarray.Find(What:="*" & SCode, After:=tbEx.ListColumns(Logx).DataBodyRange(StartE, 1), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                    On Error Resume Next
                        If Not FRow Is Nothing Then
                           YRow = FRow.Row - 1
                           'CalculDiff---------------------------------------------------------------------------------
                            CalculDiff
                           tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value
                           ValMS = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value
                           MyValMS = Format(ValMS, "0.00")
                           MyValMS = Replace(MyValMS, altDec, localDec)
                          If PORTE_M = 1 Then
                                     Set foundLabel1 = MOY.Controls(ValName(m))
                                     foundLabel1.Caption = MyValMS & CRY
                                     Set foundLabel2 = MOY.Controls(LabelName(m))
                                     foundLabel2.ForeColor = &H80000004
                                     m = m + 1
                                  Else
                                     For m = 0 To NBM
                                       Set foundLabel2 = MOY.Controls(LabelName(m))
                                       If Left(foundLabel2.Caption, 2) = monthnumber Then
                                       Set foundLabel1 = MOY.Controls(ValName(m))
                                       foundLabel1.Caption = MyValMS & CRY
                                       foundLabel2.ForeColor = &H80000004
                                       m = Right(foundLabel1.Name, 2)
                                       PORTE_M = 1
                                       Exit For
                                       End If
                                    Next m
                                 End If
                         End If
              
                       'Check Changement--------------------------------------------------------------------------------------------------------------
                                    tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 0
                                    foundLabel1.ForeColor = &H808080
              
              Next i
    
    End If

'Change Reset of only one month-------------------------------------------------------------------------------------------------------
ElseIf MO1 <> "" And MO2 = "" Then
P = 3
B = 1

        'Reset period------------------------------------
         For i = 0 To NBM
         Set foundLabel = MOY.Controls(LabelName(i))
         foundLabel.ForeColor = &H808080
         Next i

     monthdate1 = VBA.DateValue(SpecialM1 & " " & 2000)
     Starti = Month(monthdate1)
     monthnumber = Format(Starti, "00")
     SCode = monthnumber & SE
                    
    Set FRow = Sarray.Find(What:="*" & SCode, After:=tbEx.ListColumns(Logx).DataBodyRange(StartE, 1), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
        On Error Resume Next
           If Not FRow Is Nothing Then
              YRow = FRow.Row - 1
                         'CalculDiff---------------------------------------------------------------------------------
                          CalculDiff
                          tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value
                           ValMS = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value
                           MyValMS = Format(ValMS, "0.00")
                           MyValMS = Replace(MyValMS, altDec, localDec)
                          If PORTE_M = 1 Then
                                     Set foundLabel1 = MOY.Controls(ValName(m))
                                     foundLabel1.Caption = MyValMS & CRY
                                     Set foundLabel2 = MOY.Controls(LabelName(m))
                                     foundLabel2.ForeColor = &H80000004
                                     m = m + 1
                                  Else
                                     For m = 0 To NBM
                                       Set foundLabel2 = MOY.Controls(LabelName(m))
                                       If Left(foundLabel2.Caption, 2) = monthnumber Then
                                       Set foundLabel1 = MOY.Controls(ValName(m))
                                       foundLabel1.Caption = MyValMS & CRY
                                       foundLabel2.ForeColor = &H80000004
                                       m = Right(foundLabel1.Name, 2)
                                       PORTE_M = 1
                                       Exit For
                                       End If
                                    Next m
                                 End If

                                  'Check Changement--------------------------------------------------------------------------------------------------------------
                                    tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 0
                                    foundLabel1.ForeColor = &H808080
            
            End If

'Enter Reset starting at the begening-----------------------------------------------------------------------------------------------
ElseIf MO1 = "" And MO2 <> "" Then
P = 3
B = 1

        'Reset period------------------------------------
         For i = 0 To NBM
         Set foundLabel = MOY.Controls(LabelName(i))
         foundLabel.ForeColor = &H808080
         Next i

     monthdate1 = tbEx.ListColumns(Mois).DataBodyRange(StartE, 1).value
     monthdate2 = VBA.DateValue(SpecialM2 & " " & 2000)
     Starti = Month(monthdate1)
     Endi = Month(monthdate2)
          If Endi < Starti Then
             Endir = MSend
          Else
             Endir = Endi
          End If
     
            C = 0
            For i = Starti To Endir
                    If C = 1 And i = Endir + 1 Then
                       Exit For
                    ElseIf C = 1 And Endi = 12 Then
                       Exit For
                    End If
                    monthnumber = Format(i, "00")
                    SCode = monthnumber & SE
                    
                    Set FRow = Sarray.Find(What:="*" & SCode, After:=tbEx.ListColumns(Logx).DataBodyRange(StartE, 1), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
                    On Error Resume Next
                        If Not FRow Is Nothing Then
                           YRow = FRow.Row - 1
                           'CalculDiff---------------------------------------------------------------------------------
                           CalculDiff
                           tbEx.ListColumns(MontantE).DataBodyRange(YRow, 1).value = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value
                           ValMS = tbEx.ListColumns(Montant).DataBodyRange(YRow, 1).value
                           MyValMS = Replace(Format(ValMS, "0.00"), ".", Application.International(xlDecimalSeparator))
                           MyValMS = Replace(MyValMS, altDec, localDec)
                          If PORTE_M = 1 Then
                                     Set foundLabel1 = MOY.Controls(ValName(m))
                                     foundLabel1.Caption = MyValMS & CRY
                                     Set foundLabel2 = MOY.Controls(LabelName(m))
                                     foundLabel2.ForeColor = &H80000004
                                     m = m + 1
                                  Else
                                     For m = 0 To NBM
                                       Set foundLabel2 = MOY.Controls(LabelName(m))
                                       If Left(foundLabel2.Caption, 2) = monthnumber Then
                                       Set foundLabel1 = MOY.Controls(ValName(m))
                                       foundLabel1.Caption = MyValMS & CRY
                                       foundLabel2.ForeColor = &H80000004
                                       m = Right(foundLabel1.Name, 2)
                                       PORTE_M = 1
                                       Exit For
                                       End If
                                    Next m
                                 End If
                                
                         End If
                'Reset new year-----------------------------------------------------------------------------------------
                If i = 12 Then
                   Endir = Endi
                   C = 1
                   i = 0
                End If
             
                                 'Check Changement--------------------------------------------------------------------------------------------------------------
                                    tbEx.ListColumns("Check").DataBodyRange(YRow, 1).value = 0
                                    foundLabel1.ForeColor = &H808080
             
             Next i

C = 0
End If


                   'Set Reset Value-----------------------------------------------------------------------------------
                    If B = 1 Then
                    ValMS = 0
                     For i = 0 To NBM
                    Set foundLabel1 = MOY.Controls(ValName(i))
                    MyValMS = Me.foundLabel1.Caption
                    MyValMS = Replace(MyValMS, CRY, "")
                    MyValMS = Replace(MyValMS, altDec, localDec)
                    ValMSCount = CDbl(Evaluate("=VALUE(""" & MyValMS & """)"))
                        If Me.foundLabel1.Caption = "" Then
                        ValMSCount = 0
                        End If
                        ValMS = ValMS + ValMSCount
                    Next i
                    TTB = ValMS
                    tbTPC.ListColumns(Montant).DataBodyRange(SE, 1).value = TTB
                    MyTTB = Abs(TTB) & CRY
                    MyTTB = Replace(MyTTB, altDec, localDec)
                    Me.TBudget.Caption = MyTTB & CRY
                    TTDiff = TTB - TTREF
                    tbTPC.ListColumns("Diff").DataBodyRange(SE, 1).value = TTDiff
                    MBAV = TTB / MSend
                    MBAV = Format(MBAV, "0.00")
                    Me.BAV.Caption = MBAV & CRY
                    MREF = TTREF / MSend
                    MREF = Format(MREF, "0.00")
                    Me.REFAV.Caption = MREF & CRY
                      If MREF = MBAV Then
                         Me.TBudget.ForeColor = &H808080
                         Me.BAV.ForeColor = &H808080
                       Else
                         Me.TBudget.ForeColor = &H80000004
                         Me.BAV.ForeColor = &H80000004
                      End If
                    LibR = ThisWorkbook.Sheets("Library").Range("A119").value
                    Me.FrameDiff.Caption = LibR
                    LibR = ThisWorkbook.Sheets("Library").Range("A115").value
                    Me.LabelDiff.Caption = LibR
                    Me.LabelDiff.ForeColor = &H80000004
                    If TTREF = 0 Or TTREF = "" Then
                    TTP = 100
                    Else
                    TTP = ((TTB / TTREF) - 1) * 100
                    TTP = Format(TTP, "0.00")
                    End If
                    If TTDiff < 0 Then
                    LibR = ThisWorkbook.Sheets("Library").Range("A125").value
                    Me.GainTotal.Caption = LibR
                    Me.LabelTTDiff.ForeColor = &H80000004
                    Me.LabelTTDiffP.ForeColor = &H80000004
                    ElseIf TTDiff = 0 Then
                    LibR = ThisWorkbook.Sheets("Library").Range("A118").value
                    Me.GainTotal.Caption = LibR
                    Me.LabelTTDiff.ForeColor = &H808080
                    Me.LabelTTDiffP.ForeColor = &H808080
                    ElseIf TTDiff > 0 Then
                    LibR = ThisWorkbook.Sheets("Library").Range("A118").value
                    Me.GainTotal.Caption = LibR
                    Me.LabelTTDiff.ForeColor = &H80000004
                    Me.LabelTTDiffP.ForeColor = &H80000004
                    End If
                    TTDiff = Format(TTDiff, "0.00")
                    Me.LabelTTDiff.Caption = TTDiff & CRY
                    If TTP < 0 Then
                    Me.LabelTTDiffP.Caption = "(" & TTP & "%" & ")"
                    Else
                    Me.LabelTTDiffP.Caption = "(+" & TTP & "%" & ")"
                    End If
                    End If
                    Me.Average.value = ""
                    Me.Pond.value = ""
                    Me.CashFlow.Caption = Format(tbTP.ListColumns("StartE").DataBodyRange(2, 1).value, "0.00") & CRY
P = 0
Me.Average.ForeColor = vbWhite
Me.Pond.ForeColor = vbWhite
Me.Average.BorderStyle = fmBorderStyleNone
Me.Pond.BorderStyle = fmBorderStyleNone
End Sub
Private Sub UserForm_Initialize()
Dim LibR As String
SetUserFormBorderColor Me, RGB(0, 0, 0)

localDec = Application.International(xlDecimalSeparator)
    
    ' Define the alternative decimal separator: if local is dot then alt is comma; otherwise, alt is dot.
    If localDec = "." Then
        altDec = ","
    Else
        altDec = "."
    End If

vbaSep = Mid$(Format(1.1, "0.0"), 2, 1)

CodeM = ThisWorkbook.Sheets("Library").Range("R14").value
LibR = ThisWorkbook.Sheets("Library").Range("A108").value
Me.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A109").value
Me.Label26.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A110").value
Me.Label34.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A111").value
Me.Label35.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A112").value
Me.Label50.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A113").value
Me.CommandP.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A114").value
Me.CommandM.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A115").value
Me.ResetM.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A248").value
Me.Label60.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A116").value
Me.Label61.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A117").value
Me.Label62.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A118").value
Me.GainTotal.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A119").value
Me.FrameDiff.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A120").value
Me.Label57.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A121").value
Me.Label58.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A122").value
Me.Label63.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A123").value
Me.ResetAll.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A124").value
Me.Fermer.Caption = LibR

Set tbLmmmm = ThisWorkbook.Sheets("Library").ListObjects("tbLibrarymm")
Set tbMonth = ThisWorkbook.Sheets("WB_Data").ListObjects("Month")
Set tbCCY = ThisWorkbook.Sheets("WB_Data").ListObjects("tbCCY")
CRY = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
Me.SIG.Caption = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
Set tbh = ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead")
Set REF = ThisWorkbook.Sheets("WB_Data").ListObjects("Refund")
Set TbB = ThisWorkbook.Sheets("WB_Data").ListObjects("Budget")
Logx = tbh.ListColumns("Header6").DataBodyRange(1, 1).value
Année = tbh.ListColumns("Header1").DataBodyRange(1, 1).value
Mois = tbh.ListColumns("Header2").DataBodyRange(1, 1).value
Montant = tbh.ListColumns("Header7").DataBodyRange(1, 1).value
MontantE = tbh.ListColumns("Header9").DataBodyRange(1, 1).value
maxRetries = 1
If TbB.ListColumns("Input").DataBodyRange(1, 1).value = 1 Then
   BDG = TbB.ListColumns("Budget").DataBodyRange(1, 1).value
ElseIf TbB.ListColumns("Input").DataBodyRange(2, 1).value = 1 Then
   BDG = TbB.ListColumns("Budget").DataBodyRange(2, 1).value
Else
Exit Sub
End If
         Tamp = "Tampon" & BDG
         Set tbEx = ThisWorkbook.Sheets(Tamp).ListObjects("Extrap" & BDG)
         Set tbTP = ThisWorkbook.Sheets(Tamp).ListObjects("Tampon" & BDG)
         Set tbTPC = ThisWorkbook.Sheets(Tamp).ListObjects("tbTPC" & BDG)
         Set tbEM = ThisWorkbook.Sheets(Tamp).ListObjects("EMois" & BDG)
         
 Select Case BDG
        Case "Budget1"
             Me.MO_1.RowSource = "EMoisBudget1[Mois]"
             Me.MO_2.RowSource = "EMoisBudget1[Mois]"
             Me.Cat.RowSource = "TamponBudget1[Cat]"
        Case "Budget2"
             Me.MO_1.RowSource = "EMoisBudget2[Mois]"
             Me.MO_2.RowSource = "EMoisBudget2[Mois]"
             Me.Cat.RowSource = "TamponBudget2[Cat]"
End Select
         


StartE = tbTP.ListColumns("StartE").DataBodyRange(1, 1).value
EndE = tbTP.ListColumns("EndE").DataBodyRange(1, 1).value
LabelName = Array("Label_01", "Label_02", "Label_03", "Label_04", "Label_05", "Label_06", "Label_07", "Label_08", "Label_09", "Label_010", "Label_011", "Label_012")
ValName = Array("Val_01", "Val_02", "Val_03", "Val_04", "Val_05", "Val_06", "Val_07", "Val_08", "Val_09", "Val_10", "Val_11", "Val_12")
ValRName = Array("ValR_01", "ValR_02", "ValR_03", "ValR_04", "ValR_05", "ValR_06", "ValR_07", "ValR_08", "ValR_09", "ValR_10", "ValR_11", "ValR_12")
MSend = tbTP.ListColumns("NbMonth").DataBodyRange(1, 1).value
NBM = MSend - 1

  For i = 0 To 11
  Set foundLabel = MOY.Controls(LabelName(i))
  m = i + 1
  monthStart = tbEM.ListColumns("Mois").DataBodyRange(m, 1).value
     If monthStart = "" Then
        foundLabel.Caption = ""
     Else
                    Set FRow = tbMonth.ListColumns("Mois").DataBodyRange.Find(What:=monthStart, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                    On Error Resume Next
                        If Not FRow Is Nothing Then
                           YRow = FRow.Row - 1
                           monthGet = tbMonth.ListColumns(Mois).DataBodyRange(YRow, 1).value
                        End If
        monthdate1 = VBA.DateValue(YRow & " " & 2000)
        AM = Month(monthdate1)
        monthnumber = Format(YRow, "00")
        foundLabel.Caption = monthnumber & " " & monthStart & ":"
    End If
   Set foundLabel = MOY.Controls(ValName(i))
   foundLabel.Caption = ""
   Set foundLabel = MOY.Controls(ValRName(i))
   foundLabel.Caption = ""
   Next i


'Label------------------------------------------------------------------------------------
Me.Pond.value = ""
Me.LabelDiff.Caption = ""
Me.LabelTTDiff.Caption = ""
Me.LabelTTDiffP.Caption = ""
Me.TBudget.Caption = ""
Me.TREF.Caption = ""
Me.BAV.Caption = ""
Me.REFAV.Caption = ""
Me.Average.value = ""
Me.Commandquit.Caption = BDG
Me.CashFlow.Caption = Format(tbTP.ListColumns("StartE").DataBodyRange(2, 1).value, "0.00") & CRY

MS = Month(tbEx.ListColumns("Mois").DataBodyRange(StartE, 1).value)
monthStart = tbEM.ListColumns("Mois").DataBodyRange(StartE, 1).value
monthStart = monthStart & " " & tbEx.ListColumns("Année").DataBodyRange(StartE, 1).value
Mend = Month(tbEx.ListColumns("Mois").DataBodyRange(EndE, 1).value)
monthEnd = tbEM.ListColumns("Mois").DataBodyRange(MSend, 1).value
monthEnd = monthEnd & " " & tbEx.ListColumns("Année").DataBodyRange(EndE, 1).value
MOY.M_1.Caption = monthStart
MOY.M_2.Caption = monthEnd
MOY.YE1.Caption = tbEx.ListColumns(Année).DataBodyRange(EndE, 1).value
MOY.YE11.Caption = tbEx.ListColumns(Année).DataBodyRange(EndE, 1).value
If tbEx.ListColumns(Année).DataBodyRange(StartE, 1).value = tbEx.ListColumns(Année).DataBodyRange(EndE, 1).value Then
SameYear = tbEx.ListColumns(Année).DataBodyRange(StartE, 1).value
MOY.YE22.Caption = SameYear - 1
MOY.YE2.Caption = SameYear - 1
Else
MOY.YE22.Caption = tbEx.ListColumns(Année).DataBodyRange(StartE, 1).value
MOY.YE2.Caption = tbEx.ListColumns(Année).DataBodyRange(StartE, 1).value
End If

End Sub

