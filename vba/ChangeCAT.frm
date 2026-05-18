VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChangeCAT 
   Caption         =   "Modifier vos catégories"
   ClientHeight    =   6140
   ClientLeft      =   112
   ClientTop       =   448
   ClientWidth     =   7528
   OleObjectBlob   =   "ChangeCAT.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ChangeCAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cat2, Cat3TA, Cat3TB, REF, tbh, tbLm, tbLmm, tb1, tbMonth, tb2, TbEM1, TbEM2, CatMOD, TbB1, TbEx1, TbTP1, TbTPC1, TbB2, TbEx2, TbTP2, TbTPC2, tbMOF, Cat2I, Cat3TAI, Cat3TBI, tbCCY, tbRefin As ListObject
Public Row3, RowI, i, y, YRow As Long
Public FRow, RowIndex, Sarray3, Sarray2, Sarray, DefaultUser As Range
Public TransA, TransB, CatA2, CatB2, C3, C2, MODF, NewCAT, MPTASS, GCAT, GCATB, TransAI, TransBI, CatA2I, CatB2I, STRINC, CRY, LibRD, LibRI, USER As String
Public Data3, Data2 As Variant
Public A, z, ER, PINC, ValNFFC, ValNFCA, ValNFCB As Integer

Private Sub AD1_Click()
Dim reponse As VbMsgBoxResult
Dim tbcat12, tbcat11, tbcat2 As ListObject
Dim Tbrow1, Tbrow2, i As Long
Dim A As Integer
Dim dataArr1, dataArr2 As Variant
Dim Cat10, LibR As String
Dim rng11, rng12 As Range
Application.ScreenUpdating = False
ER = 0
Me.LabelCAT.Caption = ""
Me.NFCB.ForeColor = &HFFFFFF
Me.NFCA.ForeColor = &HFFFFFF
Me.NFFC.ForeColor = &HFFFFFF
          If Me.Cat2A.value = "" Then
             Me.Cat2A.BorderStyle = fmBorderStyleSingle
             Me.Cat2A.BorderColor = vbRed
          ER = 1
          End If


          If Me.AD1CAT3.value = "" Then
          ER = 1
          ElseIf Me.AD1CAT3.value = "Dépense" Or Me.AD1CAT3.value = "Income" Then
          Me.AD1CAT3.ForeColor = vbRed
          LibR = ThisWorkbook.Sheets("Library").Range("A62").value
          MsgBox (LibR), vbOK
          ER = 1
          End If
          If ER = 1 Then
          Exit Sub
          End If

Set tbcat2 = ThisWorkbook.Sheets("WB_Data").ListObjects(CatA2)
ValNFCA = tbcat2.ListRows.Count
Tbrow2 = REF.ListRows.Count
If ValNFCA = 15 Then
   Me.NFCA.ForeColor = &HFF&
   ER = 1
ElseIf Tbrow2 = 99 Then
   Me.NFFC.ForeColor = &HFF&
   ER = 1
End If
If ER = 1 Then
Exit Sub
End If

                                                Set Sarray = REF.ListColumns("C224").DataBodyRange
                                                On Error Resume Next
                                                Set FRow = Sarray.Find(What:=Me.AD1CAT3.value, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                If Not FRow Is Nothing Then
                                                LibR = ThisWorkbook.Sheets("Library").Range("A57").value
                                                MsgBox (LibR), vbExclamation
                                                Me.AD1CAT3.value = ""
                                                Exit Sub
                                                End If
                                                

   LibR = ThisWorkbook.Sheets("Library").Range("A63").value
   reponse = MsgBox((LibR), vbOKCancel)
   If reponse = vbCancel Then
   Exit Sub
   End If

CatA2 = Me.Cat2A.value
Set tbcat11 = ThisWorkbook.Sheets("WB_Data").ListObjects("Income")
Set tbcat12 = ThisWorkbook.Sheets("WB_Data").ListObjects("Dépense")
Set tbcat2 = ThisWorkbook.Sheets("WB_Data").ListObjects(CatA2)
Set REF = ThisWorkbook.Sheets("WB_Data").ListObjects("Refund")
i = 0
A = 0
Set rng11 = tbcat11.ListColumns(1).DataBodyRange
Set rng12 = tbcat12.ListColumns(1).DataBodyRange

dataArr1 = rng11.value
dataArr2 = rng12.value
    
    ' Loop through the array to process each value
    For i = LBound(dataArr2, 1) To UBound(dataArr2, 1)
            If CatA2 = dataArr2(i, 1) Then
               A = 1
               Exit For
            End If
    Next i
    i = 0
    
    If A = 1 Then
    Cat10 = "Dépense"
    Else
            For i = LBound(dataArr1, 1) To UBound(dataArr1, 1)
                    If CatA2 = dataArr1(i, 1) Then
                       Cat10 = "Income"
                       Exit For
                    End If
            Next i
    End If

Tbrow1 = tbcat2.ListRows.Count + 1
tbcat2.ListColumns(1).DataBodyRange(Tbrow1, 1).value = Me.AD1CAT3.value
REF.ListRows.Add
REF.Resize REF.Range.Resize(REF.ListRows.Count + 1, REF.ListColumns.Count)
Tbrow2 = REF.ListRows.Count
REF.ListColumns("Cat_1").DataBodyRange(Tbrow2, 1).value = Cat10
REF.ListColumns("Cat_2").DataBodyRange(Tbrow2, 1).value = CatA2
REF.ListColumns("C224").DataBodyRange(Tbrow2, 1).value = Me.AD1CAT3.value
REF.ListColumns("Row").DataBodyRange(Tbrow2, 1).value = Tbrow2
Me.NFFC.Caption = Tbrow2
Me.NFFC.ForeColor = &HFFFFFF

Set tbcat2 = ThisWorkbook.Sheets("WB_Data").ListObjects(CatA2)
ValNFCA = tbcat2.ListRows.Count
Me.NFCA.Caption = ValNFCA
Me.NFCA.ForeColor = &HFFFFFF
Me.Cat3A.RowSource = tbcat2
Me.AD1CAT3.value = ""
LibR = ThisWorkbook.Sheets("Library").Range("A61").value
Me.LabelCAT.Caption = LibR
z = 1
End Sub
Private Sub AD1I_Click()
Dim reponse As VbMsgBoxResult
Dim tbcat12, tbcat11, tbcat2 As ListObject
Dim Tbrow1, Tbrow2, i As Long
Dim A As Integer
Dim dataArr1, dataArr2 As Variant
Dim Cat10, LibR As String
Dim rng11, rng12 As Range
Application.ScreenUpdating = False
ER = 0

Me.NFFCI.ForeColor = &HFFFFFF
          If Me.Cat2AI.value = "" Then
             Me.Cat2AI.BorderStyle = fmBorderStyleSingle
             Me.Cat2AI.BorderColor = vbRed
          ER = 1
          End If

          Me.LabelCATI.Caption = ""
          
          If Me.AD1CAT3I.value = "" Then
          ER = 1
          ElseIf Me.AD1CAT3I.value = "Dépense" Or Me.AD1CAT3I.value = "Income" Then
          Me.AD1CAT3I.ForeColor = vbRed
          LibR = ThisWorkbook.Sheets("Library").Range("A62").value
          MsgBox (LibR), vbOK
          ER = 1
          End If
          If ER = 1 Then
          Exit Sub
          End If

Tbrow2 = REF.ListRows.Count
If Tbrow2 = 99 Then
   Me.NFFCI.ForeColor = &HFF&
   Exit Sub
End If

          
                                                Set Sarray = REF.ListColumns("C224").DataBodyRange
                                                On Error Resume Next
                                                Set FRow = Sarray.Find(What:=Me.AD1CAT3I.value, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                If Not FRow Is Nothing Then
                                                LibR = ThisWorkbook.Sheets("Library").Range("A57").value
                                                MsgBox (LibR), vbExclamation
                                                Me.AD1CAT3I.value = ""
                                                Exit Sub
                                                End If
                                                


   LibR = ThisWorkbook.Sheets("Library").Range("A63").value
   reponse = MsgBox((LibR), vbOKCancel)
   If reponse = vbCancel Then
   Exit Sub
   End If
CatA2I = Me.Cat2AI.value
Set tbcat11 = ThisWorkbook.Sheets("WB_Data").ListObjects("Income")
Set tbcat12 = ThisWorkbook.Sheets("WB_Data").ListObjects("Dépense")
Set tbcat2 = ThisWorkbook.Sheets("WB_Data").ListObjects(CatA2I)
Set REF = ThisWorkbook.Sheets("WB_Data").ListObjects("Refund")
i = 0
A = 0
Set rng11 = tbcat11.ListColumns(1).DataBodyRange
Set rng12 = tbcat12.ListColumns(1).DataBodyRange

dataArr1 = rng11.value
dataArr2 = rng12.value
    
    ' Loop through the array to process each value
    For i = LBound(dataArr2, 1) To UBound(dataArr2, 1)
            If CatA2I = dataArr2(i, 1) Then
               A = 1
               Exit For
            End If
    Next i
    i = 0
    
    If A = 1 Then
    Cat10 = "Dépense"
    Else
            For i = LBound(dataArr1, 1) To UBound(dataArr1, 1)
                    If CatA2I = dataArr1(i, 1) Then
                       Cat10 = "Income"
                       Exit For
                    End If
            Next i
    End If

Tbrow1 = tbcat2.ListRows.Count + 1
tbcat2.ListColumns(1).DataBodyRange(Tbrow1, 1).value = Me.AD1CAT3I.value
REF.ListRows.Add
REF.Resize REF.Range.Resize(REF.ListRows.Count + 1, REF.ListColumns.Count)
Tbrow2 = REF.ListRows.Count
REF.ListColumns("Cat_1").DataBodyRange(Tbrow2, 1).value = Cat10
REF.ListColumns("Cat_2").DataBodyRange(Tbrow2, 1).value = CatA2I
REF.ListColumns("C224").DataBodyRange(Tbrow2, 1).value = Me.AD1CAT3I.value
REF.ListColumns("Row").DataBodyRange(Tbrow2, 1).value = Tbrow2
Me.NFFCI.Caption = Tbrow2
Me.NFFCI.ForeColor = &HFFFFFF

Set tbcat2 = ThisWorkbook.Sheets("WB_Data").ListObjects(CatA2I)
Me.Cat3AI.RowSource = tbcat2
Me.AD1CAT3I.value = ""
LibR = ThisWorkbook.Sheets("Library").Range("A61").value
Me.LabelCATI.Caption = LibR
z = 1
End Sub
Private Sub AD1CAT3_Change()
Me.LabelCAT.Caption = ""
Me.AD1CAT3.ForeColor = vbWhite
End Sub
Private Sub AD1CAT3I_Change()
Me.LabelCATI.Caption = ""
Me.AD1CAT3I.ForeColor = vbWhite
End Sub
Private Sub AD2_Click()
Dim reponse As VbMsgBoxResult
Dim tbcat12, tbcat11, tbcat2 As ListObject
Dim Tbrow1, Tbrow2, i As Long
Dim A As Integer
Dim dataArr1, dataArr2 As Variant
Dim Cat10, LibR As String
Dim rng11, rng12 As Range
Application.ScreenUpdating = False
Me.LabelCAT.Caption = ""
Me.NFCB.ForeColor = &HFFFFFF
Me.NFCA.ForeColor = &HFFFFFF
Me.NFFC.ForeColor = &HFFFFFF
ER = 0
          If Me.Cat2B.value = "" Then
             Me.Cat2B.BorderStyle = fmBorderStyleSingle
             Me.Cat2B.BorderColor = vbRed
          ER = 1
          End If

          If Me.AD2CAT3.value = "" Then
          ER = 1
          ElseIf Me.AD2CAT3.value = "Dépense" Or Me.AD2CAT3.value = "Income" Then
          Me.AD2CAT3.ForeColor = vbRed
          LibR = ThisWorkbook.Sheets("Library").Range("A62").value
          MsgBox (LibR), vbOK
          ER = 1
          End If
          If ER = 1 Then
          Exit Sub
          End If
Set tbcat2 = ThisWorkbook.Sheets("WB_Data").ListObjects(CatB2)
ValNFCB = tbcat2.ListRows.Count
Tbrow2 = REF.ListRows.Count

If ValNFCB = 15 Then
   Me.NFCB.ForeColor = &HFF&
   ER = 1
ElseIf Tbrow2 = 99 Then
   Me.NFFC.ForeColor = &HFF&
   ER = 1
End If
If ER = 1 Then
Exit Sub
End If

                                                Set Sarray = REF.ListColumns("C224").DataBodyRange
                                                On Error Resume Next
                                                Set FRow = Sarray.Find(What:=Me.AD2CAT3.value, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                If Not FRow Is Nothing Then
                                                LibR = ThisWorkbook.Sheets("Library").Range("A57").value
                                                MsgBox (LibR), vbExclamation
                                                Me.AD2CAT3.value = ""
                                                Exit Sub
                                                End If
                                                



   LibR = ThisWorkbook.Sheets("Library").Range("A63").value
   reponse = MsgBox((LibR), vbOKCancel)
   If reponse = vbCancel Then
   Exit Sub
   End If
          
CatB2 = Me.Cat2B.value
Set tbcat11 = ThisWorkbook.Sheets("WB_Data").ListObjects("Income")
Set tbcat12 = ThisWorkbook.Sheets("WB_Data").ListObjects("Dépense")
Set tbcat2 = ThisWorkbook.Sheets("WB_Data").ListObjects(CatB2)
Set REF = ThisWorkbook.Sheets("WB_Data").ListObjects("Refund")
i = 0
A = 0
Set rng11 = tbcat11.ListColumns(1).DataBodyRange
Set rng12 = tbcat12.ListColumns(1).DataBodyRange

dataArr1 = rng11.value
dataArr2 = rng12.value
    
    ' Loop through the array to process each value
    For i = LBound(dataArr2, 1) To UBound(dataArr2, 1)
            If CatB2 = dataArr2(i, 1) Then
               A = 1
               Exit For
            End If
    Next i
    i = 0
    
    If A = 1 Then
    Cat10 = "Dépense"
    Else
            For i = LBound(dataArr1, 1) To UBound(dataArr1, 1)
                    If CatB2 = dataArr1(i, 1) Then
                       Cat10 = "Income"
                       Exit For
                    End If
            Next i
    End If

Tbrow1 = tbcat2.ListRows.Count + 1
tbcat2.ListColumns(1).DataBodyRange(Tbrow1, 1).value = Me.AD2CAT3.value
REF.ListRows.Add
REF.Resize REF.Range.Resize(REF.ListRows.Count + 1, REF.ListColumns.Count)
Tbrow2 = REF.ListRows.Count
REF.ListColumns("Cat_1").DataBodyRange(Tbrow2, 1).value = Cat10
REF.ListColumns("Cat_2").DataBodyRange(Tbrow2, 1).value = CatB2
REF.ListColumns("C224").DataBodyRange(Tbrow2, 1).value = Me.AD2CAT3.value
REF.ListColumns("Row").DataBodyRange(Tbrow2, 1).value = Tbrow2
Me.NFFC.Caption = Tbrow2
Me.NFFC.ForeColor = &HFFFFFF
Set tbcat2 = ThisWorkbook.Sheets("WB_Data").ListObjects(CatB2)
ValNFCB = tbcat2.ListRows.Count
Me.NFCB.Caption = ValNFCB
Me.NFCB.ForeColor = &HFFFFFF
Me.Cat3B.RowSource = tbcat2
Me.AD2CAT3.value = ""
LibR = ThisWorkbook.Sheets("Library").Range("A61").value
Me.LabelCAT.Caption = LibR
z = 1
End Sub
Private Sub AD2I_Click()
Dim reponse As VbMsgBoxResult
Dim tbcat12, tbcat11, tbcat2 As ListObject
Dim Tbrow1, Tbrow2, i As Long
Dim A As Integer
Dim dataArr1, dataArr2 As Variant
Dim Cat10, LibR As String
Dim rng11, rng12 As Range
Application.ScreenUpdating = False
Me.LabelCATI.Caption = ""
ER = 0

          If Me.Cat2BI.value = "" Then
             Me.Cat2BI.BorderStyle = fmBorderStyleSingle
             Me.Cat2BI.BorderColor = vbRed
          ER = 1
          End If

          If Me.AD2CAT3I.value = "" Then
          ER = 1
          ElseIf Me.AD2CAT3I.value = "Dépense" Or Me.AD2CAT3I.value = "Income" Then
          Me.AD2CAT3I.ForeColor = vbRed
          LibR = ThisWorkbook.Sheets("Library").Range("A62").value
          MsgBox (LibR), vbOK
          ER = 1
          End If
          If ER = 1 Then
          Exit Sub
          End If
Tbrow2 = REF.ListRows.Count
If Tbrow2 = 99 Then
   Me.NFFCI.ForeColor = &HFF&
   Exit Sub
End If
                                                Set Sarray = REF.ListColumns("C224").DataBodyRange
                                                On Error Resume Next
                                                Set FRow = Sarray.Find(What:=Me.AD2CAT3I.value, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                If Not FRow Is Nothing Then
                                                LibR = ThisWorkbook.Sheets("Library").Range("A57").value
                                                MsgBox (LibR), vbExclamation
                                                Me.AD1CAT3I.value = ""
                                                Exit Sub
                                                End If
                                                


   LibR = ThisWorkbook.Sheets("Library").Range("A63").value
   reponse = MsgBox((LibR), vbOKCancel)
   If reponse = vbCancel Then
   Exit Sub
   End If
CatB2I = Me.Cat2BI.value
Set tbcat11 = ThisWorkbook.Sheets("WB_Data").ListObjects("Income")
Set tbcat12 = ThisWorkbook.Sheets("WB_Data").ListObjects("Dépense")
Set tbcat2 = ThisWorkbook.Sheets("WB_Data").ListObjects(CatB2I)
Set REF = ThisWorkbook.Sheets("WB_Data").ListObjects("Refund")
i = 0
A = 0
Set rng11 = tbcat11.ListColumns(1).DataBodyRange
Set rng12 = tbcat12.ListColumns(1).DataBodyRange

dataArr1 = rng11.value
dataArr2 = rng12.value
    
    ' Loop through the array to process each value
    For i = LBound(dataArr2, 1) To UBound(dataArr2, 1)
            If CatB2I = dataArr2(i, 1) Then
               A = 1
               Exit For
            End If
    Next i
    i = 0
    
    If A = 1 Then
    Cat10 = "Dépense"
    Else
            For i = LBound(dataArr1, 1) To UBound(dataArr1, 1)
                    If CatB2I = dataArr1(i, 1) Then
                       Cat10 = "Income"
                       Exit For
                    End If
            Next i
    End If

Tbrow1 = tbcat2.ListRows.Count + 1
tbcat2.ListColumns(1).DataBodyRange(Tbrow1, 1).value = Me.AD2CAT3I.value
REF.ListRows.Add
REF.Resize REF.Range.Resize(REF.ListRows.Count + 1, REF.ListColumns.Count)
Tbrow2 = REF.ListRows.Count
REF.ListColumns("Cat_1").DataBodyRange(Tbrow2, 1).value = Cat10
REF.ListColumns("Cat_2").DataBodyRange(Tbrow2, 1).value = CatB2I
REF.ListColumns("C224").DataBodyRange(Tbrow2, 1).value = Me.AD2CAT3I.value
REF.ListColumns("Row").DataBodyRange(Tbrow2, 1).value = Tbrow2
Me.NFFCI.Caption = Tbrow2
Me.NFFCI.ForeColor = &HFFFFFF
Set tbcat2 = ThisWorkbook.Sheets("WB_Data").ListObjects(CatB2I)
Me.Cat3BI.RowSource = tbcat2
Me.AD2CAT3I.value = ""
LibR = ThisWorkbook.Sheets("Library").Range("A61").value
Me.LabelCATI.Caption = LibR
z = 1
End Sub
Private Sub AD2CAT3_Change()
Me.LabelCAT.Caption = ""
Me.AD2CAT3.ForeColor = vbWhite
End Sub
Private Sub AD2CAT3I_Change()
Me.LabelCATI.Caption = ""
Me.AD2CAT3I.ForeColor = vbWhite
End Sub
Private Sub Cat2A_Change()
Me.Cat2A.BorderStyle = fmBorderStyleNone
Me.LabelCAT.Caption = ""
If Cat2A.value <> "" Then

   If Cat2A.value = Cat2B.value Then
      Cat2A.value = ""
      Me.Cat3A.RowSource = ""
      Me.NFCA.Caption = ""
      Me.NFCA.ForeColor = &HFFFFFF
      Me.NFA.Visible = False
      Exit Sub
   End If
Me.Cat3A.value = ""
CatA2 = Me.Cat2A.value
Set Cat3TA = ThisWorkbook.Sheets("WB_Data").ListObjects(CatA2)
Me.Cat3A.RowSource = Cat3TA
ValNFCA = ThisWorkbook.Sheets("WB_Data").ListObjects(CatA2).ListRows.Count
Me.NFCA.Caption = ValNFCA
Me.NFCA.ForeColor = &HFFFFFF
Me.NFA.Visible = True
End If
End Sub

Private Sub Cat2B_Change()
Me.Cat2B.BorderStyle = fmBorderStyleNone
Me.LabelCAT.Caption = ""
If Cat2B.value <> "" Then

   If Cat2B.value = Cat2A.value Then
      Cat2B.value = ""
      Me.Cat3B.RowSource = ""
      Me.NFCB.Caption = ""
      Me.NFCB.ForeColor = &HFFFFFF
      Me.NFB.Visible = False
      Exit Sub
   End If

Me.Cat3B.value = ""
CatB2 = Me.Cat2B.value
Set Cat3TB = ThisWorkbook.Sheets("WB_Data").ListObjects(CatB2)
Me.Cat3B.RowSource = Cat3TB
ValNFCB = ThisWorkbook.Sheets("WB_Data").ListObjects(CatB2).ListRows.Count
Me.NFCB.Caption = ValNFCB
Me.NFCB.ForeColor = &HFFFFFF
Me.NFB.Visible = True
End If
End Sub
Private Sub Cat2AI_Change()
Me.LabelCATI.Caption = ""

Me.Cat2AI.BorderStyle = fmBorderStyleNone

If Cat2AI.value <> "" Then

   If Cat2AI.value = Cat2BI.value Then
      Cat2AI.value = ""
      Me.Cat3AI.RowSource = ""
      Exit Sub
   End If
Me.Cat3AI.value = ""
CatA2I = Me.Cat2AI.value
Set Cat3TAI = ThisWorkbook.Sheets("WB_Data").ListObjects(CatA2I)
Me.Cat3AI.RowSource = Cat3TAI
End If
End Sub

Private Sub Cat2BI_Change()
Me.LabelCATI.Caption = ""

Me.Cat2BI.BorderStyle = fmBorderStyleNone

If Cat2BI.value <> "" Then

   If Cat2BI.value = Cat2AI.value Then
      Cat2BI.value = ""
      Me.Cat3BI.RowSource = ""
      Exit Sub
   End If

Me.Cat3BI.value = ""
CatB2I = Me.Cat2BI.value
Set Cat3TBI = ThisWorkbook.Sheets("WB_Data").ListObjects(CatB2I)
Me.Cat3BI.RowSource = Cat3TBI
End If
End Sub
Private Sub Cat3A_Click()
Me.Cat3B.value = Null
End Sub

Private Sub Cat3B_Click()
Me.Cat3A.value = Null
End Sub
Private Sub Cat3AI_Click()
Me.Cat3BI.value = Null
End Sub

Private Sub Cat3BI_Click()
Me.Cat3AI.value = Null
End Sub

Private Sub ChoiceL_Change()
Me.ChoiceL.BorderStyle = fmBorderStyleNone
Me.ChoiceL.ForeColor = vbWhite
LibR = ThisWorkbook.Sheets("Library").Range("A215").value
Me.Frame8.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A217").value
Me.Frame9.Caption = LibR
End Sub

Private Sub Comandquit_Click()
Dim Category_1, Category_2, Category_3, Category_4, Category_5, Category_6, Category_7  As String
Dim tbDP As ListObject
Dim pt As PivotTable

 If z = 1 Then
    Application.ScreenUpdating = False
    ThisWorkbook.Sheets("Master").ListObjects("Master_1").QueryTable.Refresh BackgroundQuery:=False

    ' Refresh all
    Do While Application.CalculationState <> xlDone
    DoEvents
    Loop
    
    ThisWorkbook.RefreshAll
    
    Do While Application.CalculationState <> xlDone
    DoEvents
    Loop
    
    Set pt = ThisWorkbook.Sheets("Analyse").PivotTables("PivotTable1")
    pt.RefreshTable
    
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Set tbDP = ThisWorkbook.Sheets("WB_Data").ListObjects("Dépense")
    Category_1 = tbDP.ListColumns("C21").DataBodyRange(1, 1).value
    Category_2 = tbDP.ListColumns("C21").DataBodyRange(2, 1).value
    Category_3 = tbDP.ListColumns("C21").DataBodyRange(3, 1).value
    Category_4 = tbDP.ListColumns("C21").DataBodyRange(4, 1).value
    Category_5 = tbDP.ListColumns("C21").DataBodyRange(5, 1).value
    Category_6 = tbDP.ListColumns("C21").DataBodyRange(6, 1).value
    Category_7 = tbDP.ListColumns("C21").DataBodyRange(7, 1).value

    ThisWorkbook.Sheets("A_Budget1").PivotTables("Budget1_QA1").PivotFields( _
        "[Budget1Data].[Cat_2].[Cat_2]").VisibleItemsList = Array( _
        "[Budget1Data].[Cat_2].&[" & Category_1 & "]", "[Budget1Data].[Cat_2].&[" & Category_2 & "]", _
        "[Budget1Data].[Cat_2].&[" & Category_3 & "]", "[Budget1Data].[Cat_2].&[" & Category_4 & "]", _
        "[Budget1Data].[Cat_2].&[" & Category_5 & "]", "[Budget1Data].[Cat_2].&[" & Category_6 & "]", _
        "[Budget1Data].[Cat_2].&[" & Category_7 & "]")

   ThisWorkbook.Sheets("A_Budget2").PivotTables("Budget2_QA2").PivotFields( _
        "[DataBudget2].[Cat_2].[Cat_2]").VisibleItemsList = Array( _
        "[DataBudget2].[Cat_2].&[" & Category_1 & "]", "[DataBudget2].[Cat_2].&[" & Category_2 & "]", _
        "[DataBudget2].[Cat_2].&[" & Category_3 & "]", "[DataBudget2].[Cat_2].&[" & Category_4 & "]", _
        "[DataBudget2].[Cat_2].&[" & Category_5 & "]", "[DataBudget2].[Cat_2].&[" & Category_6 & "]", _
        "[DataBudget2].[Cat_2].&[" & Category_7 & "]")

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
 End If

 Unload Me
End Sub



Private Sub CommandButton1_Click()

End Sub

Private Sub Frame8_Click()

End Sub

Private Sub GCAT2_Change()
Me.GComplete.Caption = ""
Me.GCAT2.BorderStyle = fmBorderStyleNone
End Sub

Private Sub GCAT2I_Change()

PINC = 0
STRINC = ""
Me.GCompleteI.Caption = ""
Me.GCAT2I.BorderStyle = fmBorderStyleNone
Select Case Me.GCAT2I.value
       Case Is = "Income_1"
       PINC = 1
       STRINC = Cat2I.ListColumns("C22").DataBodyRange(1, 1).value
       If Left(STRINC, 3) = "i1_" Then
       STRINC = Mid(STRINC, 4)
       Me.LabelIn1.ForeColor = &HFFFFC0
       Me.LabelIn2.ForeColor = &H808080
       Me.LabelIn3.ForeColor = &H808080
       End If
       Me.GCAT2MI.value = STRINC
       Case Is = "Income_2"
       PINC = 2
       STRINC = Cat2I.ListColumns("C22").DataBodyRange(2, 1).value
       If Left(STRINC, 3) = "i2_" Then
       STRINC = Mid(STRINC, 4)
       Me.LabelIn1.ForeColor = &H808080
       Me.LabelIn2.ForeColor = &HFFFFC0
       Me.LabelIn3.ForeColor = &H808080
       End If
       Me.GCAT2MI.value = STRINC
       Case Is = "Income_3"
       PINC = 3
       STRINC = Cat2I.ListColumns("C22").DataBodyRange(3, 1).value
       If Left(STRINC, 3) = "i3_" Then
       STRINC = Mid(STRINC, 4)
       Me.LabelIn1.ForeColor = &H808080
       Me.LabelIn2.ForeColor = &H808080
       Me.LabelIn3.ForeColor = &HFFFFC0
       End If
       Me.GCAT2MI.value = STRINC
       Case Is = ""
       Me.LabelIn1.ForeColor = &H808080
       Me.LabelIn2.ForeColor = &H808080
       Me.LabelIn3.ForeColor = &H808080
       Me.GCAT2MI.value = ""
End Select

End Sub

Private Sub GCAT2M_Change()
Me.GComplete.Caption = ""
Me.GCAT2M.BorderStyle = fmBorderStyleNone
Me.GCAT2M.ForeColor = vbWhite
End Sub

Private Sub GCAT2MI_Change()
Me.GCompleteI.Caption = ""
Me.GCAT2MI.BorderStyle = fmBorderStyleNone
Me.GCAT2MI.ForeColor = vbWhite
End Sub

Private Sub Label9_Click()

End Sub

Private Sub LFolder_Change()
If DisableEvents Then Exit Sub
Me.LFolder.BorderStyle = fmBorderStyleNone
LibR = ThisWorkbook.Sheets("Library").Range("A217").value
Me.Frame9.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A215").value
Me.Frame8.Caption = LibR
End Sub

Private Sub LPath_Change()
If DisableEvents Then Exit Sub
Me.LPath.BorderStyle = fmBorderStyleNone
LibR = ThisWorkbook.Sheets("Library").Range("A217").value
Me.Frame9.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A215").value
Me.Frame8.Caption = LibR
End Sub

Private Sub ModifCat_Click()
Dim LibR As String
Dim TestingCat3 As PivotTable
Dim pfTestingCat3 As PivotField
Dim pi As PivotItem
Me.LabelCAT.Caption = ""

          If Me.Cat2A.value = "" Then
             Me.Cat2A.BorderStyle = fmBorderStyleSingle
             Me.Cat2A.BorderColor = vbRed
          Exit Sub
          End If

If Me.Cat3A.ListIndex = -1 Then
   Exit Sub
End If

NewCAT = Me.AD1CAT3.value
If Me.AD1CAT3.value <> "" Then
   If Me.AD1CAT3.value = "Dépense" Or Me.AD1CAT3.value = "Income" Then
   Me.AD1CAT3.ForeColor = vbRed
   LibR = ThisWorkbook.Sheets("Library").Range("A62").value
   MsgBox (LibR), vbOK
   Exit Sub
   End If
   
   
                                                Set Sarray = REF.ListColumns("C224").DataBodyRange
                                                On Error Resume Next
                                                Set FRow = Sarray.Find(What:=Me.AD1CAT3.value, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                If Not FRow Is Nothing Then
                                                LibR = ThisWorkbook.Sheets("Library").Range("A57").value
                                                MsgBox (LibR), vbExclamation
                                                Me.AD1CAT3.value = ""
                                                Exit Sub
                                                End If

   If Me.Cat3A.value = "" Or Me.Cat3A.value = Null Then
      LibR = ThisWorkbook.Sheets("Library").Range("A64").value
      Me.LabelCAT.Caption = LibR
      Me.LabelCAT.ForeColor = vbRed
      Me.AD1CAT3.value = ""
      Exit Sub
   Else
         MODF = Me.Cat3A.value
         
         Cat22 = Me.Cat2A.value
         Set CatMOD = ThisWorkbook.Sheets("WB_Data").ListObjects(CatA2)
   End If
   
   Set Sarray = CatMOD.ListColumns(1).DataBodyRange
   Set FRow = Sarray.Find(What:=NewCAT, LookIn:=xlValues, LookAt:=xlWhole)
               On Error Resume Next
               If Not FRow Is Nothing Then
                  LibR = ThisWorkbook.Sheets("Library").Range("A57").value
                  MsgBox (LibR), vbOK
                  Me.AD1CAT3.value = ""
                  Exit Sub
               End If

   Set Sarray = REF.ListColumns("C224").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = CatMOD.ListColumns(1).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = tb2.ListColumns(C3).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbB1.ListColumns(C3).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbEx1.ListColumns(C3).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbTP1.ListColumns("Cat").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbTPC1.ListColumns("Cat").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbB2.ListColumns(C3).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbEx2.ListColumns(C3).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbTP2.ListColumns("Cat").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbTPC2.ListColumns("Cat").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
    NewCAT = "[Master].[Cat_3].&[" & NewCAT & "]"
    Set Cat3TA = ThisWorkbook.Sheets("WB_Data").ListObjects(CatA2)
    Me.Cat3A.RowSource = Cat3TA
    Me.AD1CAT3.value = ""
    LibR = ThisWorkbook.Sheets("Library").Range("A61").value
    Me.LabelCAT.Caption = LibR
    Me.LabelCAT.ForeColor = vbWhite
    z = 1
    
End If
End Sub

Private Sub ModifCatB_Click()
Dim LibR As String
Me.LabelCAT.Caption = ""


          If Me.Cat2B.value = "" Then
             Me.Cat2B.BorderStyle = fmBorderStyleSingle
             Me.Cat2B.BorderColor = vbRed
          Exit Sub
          End If

If Me.Cat3B.ListIndex = -1 Then
   Exit Sub
End If

NewCAT = Me.AD2CAT3.value
If Me.AD2CAT3.value <> "" Then

   If Me.AD2CAT3.value = "Dépense" Or Me.AD2CAT3.value = "Income" Then
   Me.AD2CAT3.ForeColor = vbRed
   LibR = ThisWorkbook.Sheets("Library").Range("A62").value
   MsgBox (LibR), vbOK
   Exit Sub
   End If
   
                                                Set Sarray = REF.ListColumns("C224").DataBodyRange
                                                On Error Resume Next
                                                Set FRow = Sarray.Find(What:=Me.AD2CAT3.value, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                If Not FRow Is Nothing Then
                                                LibR = ThisWorkbook.Sheets("Library").Range("A57").value
                                                MsgBox (LibR), vbExclamation
                                                Me.AD2CAT3.value = ""
                                                Exit Sub
                                                End If
   

   If Me.Cat3B.value = "" Or Me.Cat3B.value = Null Then
      Me.LabelCAT.Caption = ""
      Me.LabelCAT.ForeColor = vbRed
      Me.AD2CAT3.value = ""
      Exit Sub
   Else
         MODF = Me.Cat3B.value
         
         Cat22 = Me.Cat2B.value
         Set CatMOD = ThisWorkbook.Sheets("WB_Data").ListObjects(CatB2)
   End If
   
   Set Sarray = CatMOD.ListColumns(1).DataBodyRange
   Set FRow = Sarray.Find(What:=NewCAT, LookIn:=xlValues, LookAt:=xlWhole)
               On Error Resume Next
               If Not FRow Is Nothing Then
                  LibR = ThisWorkbook.Sheets("Library").Range("A57").value
                  MsgBox (LibR), vbOK
                  Me.AD2CAT3.value = ""
                  Exit Sub
               End If
   
   Set Sarray = REF.ListColumns("C224").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = CatMOD.ListColumns(1).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = tb2.ListColumns(C3).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbB1.ListColumns(C3).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbEx1.ListColumns(C3).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbTP1.ListColumns("Cat").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbTPC1.ListColumns("Cat").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbB2.ListColumns(C3).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbEx2.ListColumns(C3).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbTP2.ListColumns("Cat").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbTPC2.ListColumns("Cat").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   
    Set Cat3TB = ThisWorkbook.Sheets("WB_Data").ListObjects(CatB2)
    Me.Cat3B.RowSource = Cat3TB
    Me.AD2CAT3.value = ""
    LibR = ThisWorkbook.Sheets("Library").Range("A61").value
    Me.LabelCAT.Caption = LibR
    Me.LabelCAT.ForeColor = vbWhite
    z = 1
   
End If
End Sub
Private Sub ModifCatI_Click()
Dim LibR As String
Me.LabelCATI.Caption = ""


          If Me.Cat2AI.value = "" Then
             Me.Cat2AI.BorderStyle = fmBorderStyleSingle
             Me.Cat2AI.BorderColor = vbRed
          Exit Sub
          End If

If Me.Cat3AI.ListIndex = -1 Then
   Exit Sub
End If

NewCAT = Me.AD1CAT3I.value
If Me.AD1CAT3I.value <> "" Then

   If Me.AD1CAT3I.value = "Dépense" Or Me.AD1CAT3I.value = "Income" Then
   Me.AD1CAT3I.ForeColor = vbRed
   LibR = ThisWorkbook.Sheets("Library").Range("A62").value
   MsgBox (LibR), vbOK
   Exit Sub
   End If

   
                                                Set Sarray = REF.ListColumns("C224").DataBodyRange
                                                On Error Resume Next
                                                Set FRow = Sarray.Find(What:=Me.AD1CAT3I.value, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                If Not FRow Is Nothing Then
                                                LibR = ThisWorkbook.Sheets("Library").Range("A57").value
                                                MsgBox (LibR), vbExclamation
                                                Me.AD1CAT3I.value = ""
                                                Exit Sub
                                                End If
   
   If Me.Cat3AI.value = "" Or Me.Cat3AI.value = Null Then
      LibR = ThisWorkbook.Sheets("Library").Range("A64").value
      Me.LabelCATI.Caption = LibR
      Me.LabelCATI.ForeColor = vbRed
      Me.AD1CAT3I.value = ""
      Exit Sub
   Else
         MODF = Me.Cat3AI.value
         Cat22 = Me.Cat2AI.value
         Set CatMOD = ThisWorkbook.Sheets("WB_Data").ListObjects(CatA2I)
   End If
   
   Set Sarray = CatMOD.ListColumns(1).DataBodyRange
   Set FRow = Sarray.Find(What:=NewCAT, LookIn:=xlValues, LookAt:=xlWhole)
               On Error Resume Next
               If Not FRow Is Nothing Then
                  LibR = ThisWorkbook.Sheets("Library").Range("A57").value
                  MsgBox (LibR), vbOK
                  Me.AD1CAT3I.value = ""
                  Exit Sub
               End If
   
   Set Sarray = REF.ListColumns("C224").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = CatMOD.ListColumns(1).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = tb2.ListColumns(C3).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbB1.ListColumns(C3).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbEx1.ListColumns(C3).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbTP1.ListColumns("Cat").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbTPC1.ListColumns("Cat").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbB2.ListColumns(C3).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbEx2.ListColumns(C3).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbTP2.ListColumns("Cat").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbTPC2.ListColumns("Cat").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   
    Set Cat3TAI = ThisWorkbook.Sheets("WB_Data").ListObjects(CatA2I)
    Me.Cat3AI.RowSource = Cat3TAI
    Me.AD1CAT3I.value = ""
    LibR = ThisWorkbook.Sheets("Library").Range("A61").value
    Me.LabelCATI.Caption = LibR
    Me.LabelCATI.ForeColor = vbWhite
    z = 1
   
End If
End Sub


Private Sub ModifCatBI_Click()
Dim LibR As String
Me.LabelCATI.Caption = ""

          If Me.Cat2BI.value = "" Then
             Me.Cat2BI.BorderStyle = fmBorderStyleSingle
             Me.Cat2BI.BorderColor = vbRed
          Exit Sub
          End If

If Me.Cat3BI.ListIndex = -1 Then
   Exit Sub
End If


NewCAT = Me.AD2CAT3I.value
If Me.AD2CAT3I.value <> "" Then

   If Me.AD2CAT3I.value = "Dépense" Or Me.AD2CAT3I.value = "Income" Then
   Me.AD2CAT3I.ForeColor = vbRed
   LibR = ThisWorkbook.Sheets("Library").Range("A62").value
   MsgBox (LibR), vbOK
   Exit Sub
   End If
   
                                                Set Sarray = REF.ListColumns("C224").DataBodyRange
                                                On Error Resume Next
                                                Set FRow = Sarray.Find(What:=Me.AD2CAT3I.value, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                If Not FRow Is Nothing Then
                                                LibR = ThisWorkbook.Sheets("Library").Range("A57").value
                                                MsgBox (LibR), vbExclamation
                                                Me.AD2CAT3I.value = ""
                                                Exit Sub
                                                End If
   
   
   If Me.Cat3BI.value = "" Or Me.Cat3BI.value = Null Then
      Me.LabelCATI.Caption = ""
      Me.LabelCATI.ForeColor = vbRed
      Me.AD2CAT3I.value = ""
      Exit Sub
   Else
         MODF = Me.Cat3BI.value
         Cat22 = Me.Cat2BI.value
         Set CatMOD = ThisWorkbook.Sheets("WB_Data").ListObjects(CatB2I)
   End If
   
   Set Sarray = CatMOD.ListColumns(1).DataBodyRange
   Set FRow = Sarray.Find(What:=NewCAT, LookIn:=xlValues, LookAt:=xlWhole)
               On Error Resume Next
               If Not FRow Is Nothing Then
                  LibR = ThisWorkbook.Sheets("Library").Range("A57").value
                  MsgBox (LibR), vbOK
                  Me.AD2CAT3I.value = ""
                  Exit Sub
               End If
   
   Set Sarray = REF.ListColumns("C224").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = CatMOD.ListColumns(1).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = tb2.ListColumns(C3).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbB1.ListColumns(C3).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbEx1.ListColumns(C3).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbTP1.ListColumns("Cat").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbTPC1.ListColumns("Cat").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbB2.ListColumns(C3).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbEx2.ListColumns(C3).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbTP2.ListColumns("Cat").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   Set Sarray = TbTPC2.ListColumns("Cat").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=NewCAT, LookAt:=xlWhole
   
    Set Cat3TBI = ThisWorkbook.Sheets("WB_Data").ListObjects(CatB2I)
    Me.Cat3BI.RowSource = Cat3TBI
    Me.AD2CAT3I.value = ""
    LibR = ThisWorkbook.Sheets("Library").Range("A61").value
    Me.LabelCATI.Caption = LibR
    Me.LabelCATI.ForeColor = vbWhite
    z = 1
   
End If
End Sub

Private Sub MPASS_Change()
Me.LabelMP.ForeColor = &HC0C0C0
End Sub

Private Sub MPASSI_Change()
Me.LabelMPI.ForeColor = &HC0C0C0
End Sub

Private Sub MultiPage1_Change()
Me.NFCB.ForeColor = &HFFFFFF
Me.NFCA.ForeColor = &HFFFFFF
Me.NFFC.ForeColor = &HFFFFFF
Me.NFFCI.ForeColor = &HFFFFFF
LibR = ThisWorkbook.Sheets("Library").Range("A215").value
Me.Frame8.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A217").value
Me.Frame9.Caption = LibR
If Me.MultiPage1.value = 0 Then
If Me.Cat2A.value = "" Then
   Me.NFCA.Caption = ""
   Me.NFA.Visible = False
End If
If Me.Cat2B.value = "" Then
   Me.NFCB.Caption = ""
   Me.NFB.Visible = False
End If
Me.LabelD.BackColor = &HFFFFC0
Me.LabelD.ForeColor = vbBlack
Me.LabelD.BorderStyle = fmBorderStyleNone
Me.LabelI.BackColor = vbBlack
Me.LabelI.ForeColor = &HFFFFC0
Me.LabelI.BorderStyle = fmBorderStyleSingle
ElseIf Me.MultiPage1.value = 1 Then
Me.LabelD.BackColor = &HFFFFC0
Me.LabelD.ForeColor = vbBlack
Me.LabelD.BorderStyle = fmBorderStyleNone
Me.LabelI.BackColor = vbBlack
Me.LabelI.ForeColor = &HFFFFC0
Me.LabelI.BorderStyle = fmBorderStyleSingle
ElseIf Me.MultiPage1.value = 2 Then
Me.LabelI.BackColor = &HFFFFC0
Me.LabelI.ForeColor = vbBlack
Me.LabelI.BorderStyle = fmBorderStyleNone
Me.LabelD.BackColor = vbBlack
Me.LabelD.ForeColor = &HFFFFC0
Me.LabelD.BorderStyle = fmBorderStyleSingle
ElseIf Me.MultiPage1.value = 3 Then
Me.LabelI.BackColor = &HFFFFC0
Me.LabelI.ForeColor = vbBlack
Me.LabelI.BorderStyle = fmBorderStyleNone
Me.LabelD.BackColor = vbBlack
Me.LabelD.ForeColor = &HFFFFC0
Me.LabelD.BorderStyle = fmBorderStyleSingle
ElseIf Me.MultiPage1.value = 4 Then
Me.LabelD.BackColor = vbBlack
Me.LabelD.ForeColor = &HFFFFC0
Me.LabelD.BorderStyle = fmBorderStyleSingle
Me.LabelI.BackColor = vbBlack
Me.LabelI.ForeColor = &HFFFFC0
Me.LabelI.BorderStyle = fmBorderStyleSingle
End If
End Sub

Private Sub TransfertA_Click()
Dim reponse As VbMsgBoxResult
Dim Check As Long
Dim LibR, LibR1, LibR2 As String
Check = 0
Me.LabelCAT.Caption = ""

If Me.Cat3B.ListIndex = -1 Then
   Exit Sub
End If

If Me.Cat2A.value = "" Or Me.Cat2B.value = "" Then
   Exit Sub
   Else
   Check = Cat3TB.ListRows.Count
         If Check < 2 Then
         LibR = ThisWorkbook.Sheets("Library").Range("A65").value
         MsgBox (LibR), vbOK
         Exit Sub
         End If
   If A = 0 Then
   LibR1 = ThisWorkbook.Sheets("Library").Range("A66").value
   LibR2 = ThisWorkbook.Sheets("Library").Range("A67").value
   reponse = MsgBox((LibR1 & vbCrLf & vbCrLf & LibR2), vbOKCancel)
   If reponse = vbOK Then
   A = 1
   Else
   Exit Sub
   End If
   End If


If ValNFCA = 15 Then
   Me.NFCA.ForeColor = &HFF&
   Exit Sub
End If
   
   TransA = Me.Cat3B.value

                                                Set Sarray = Cat3TA.ListColumns(1).DataBodyRange
                                                 On Error Resume Next
                                                Set RowIndex = Sarray.Find(What:=TransA, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                If Not RowIndex Is Nothing Then
                                                       LibR = ThisWorkbook.Sheets("Library").Range("A57").value
                                                       MsgBox (LibR), vbOK
                                                       Exit Sub
                                                End If
                                                
   Set Sarray = REF.ListColumns("C224").DataBodyRange
                                                On Error Resume Next
                                                Set FRow = Sarray.Find(What:=TransA, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                If Not FRow Is Nothing Then
                                                        SCode = FRow.Row - 1
                                                        'Cahnge Value---------------------------------------------------------------------------
                                                        REF.ListColumns(C2).DataBodyRange(SCode, 1).value = CatA2
                                                Else
                                                   LibR = ThisWorkbook.Sheets("Library").Range("A59").value
                                                   MsgBox (LibR), vbOK
                                                   Exit Sub
                                                End If

   Row3 = Cat3TA.ListRows.Count
   Cat3TA.ListRows.Add
   Cat3TA.Resize Cat3TA.Range.Resize(Cat3TA.ListRows.Count + 1, Cat3TA.ListColumns.Count)
   Cat3TA.DataBodyRange(Row3 + 1, 1).value = TransA
   Set Sarray = Cat3TB.ListColumns(1).DataBodyRange
                                                 On Error Resume Next
                                                Set RowIndex = Sarray.Find(What:=TransA, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                If Not RowIndex Is Nothing Then
                                                        RowI = RowIndex.Row - 1
                                                        'Cahnge Value---------------------------------------------------------------------------
                                                        Cat3TB.ListRows(RowI).Delete
                                                End If
   Set Sarray3 = tb2.ListColumns(C3).DataBodyRange
   Set Sarray2 = tb2.ListColumns(C2).DataBodyRange
   Data3 = Sarray3.value
   Data2 = Sarray2.value
    
    ' Loop through the array to check and update values
    For i = 1 To UBound(Data3, 1)
        If Data3(i, 1) = TransA Then
            ' Update the value in column A
            Sarray2(i, 1).value = CatA2 ' Replace with your desired value
        End If
    Next i
    Set Cat3TA = ThisWorkbook.Sheets("WB_Data").ListObjects(CatA2)
    Set Cat3TB = ThisWorkbook.Sheets("WB_Data").ListObjects(CatB2)
    Me.Cat3A.RowSource = Cat3TA
    Me.Cat3B.RowSource = Cat3TB
    z = 1
   End If

Me.NFCB.Caption = ValNFCB
Me.NFCA.Caption = ValNFCA
Me.NFCA.ForeColor = &HFFFFFF
Me.NFCB.ForeColor = &HFFFFFF
End Sub

Private Sub TransfertB_Click()
Dim reponse As VbMsgBoxResult
Dim Check As Long
Dim LibR, LibR1, LibR2 As String
Me.LabelCAT.Caption = ""

If Me.Cat3A.ListIndex = -1 Then
   Exit Sub
End If
Check = 0
If Me.Cat2A.value = "" Or Me.Cat2B.value = "" Then
   Exit Sub
   Else
         Check = Cat3TA.ListRows.Count
         If Check < 2 Then
         LibR = ThisWorkbook.Sheets("Library").Range("A65").value
         MsgBox (LibR), vbOK
         Exit Sub
         End If
   If A = 0 Then
   LibR1 = ThisWorkbook.Sheets("Library").Range("A66").value
   LibR2 = ThisWorkbook.Sheets("Library").Range("A67").value
   reponse = MsgBox((LibR1 & vbCrLf & vbCrLf & LibR2), vbOKCancel)
   If reponse = vbOK Then
   A = 1
   Else
   Exit Sub
   End If
   End If
If ValNFCB = 15 Then
   Me.NFCB.ForeColor = &HFF&
   Exit Sub
End If
   TransB = Me.Cat3A.value
                                                Set Sarray = Cat3TB.ListColumns(1).DataBodyRange
                                                 On Error Resume Next
                                                Set RowIndex = Sarray.Find(What:=TransB, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                If Not RowIndex Is Nothing Then
                                                       LibR = ThisWorkbook.Sheets("Library").Range("A57").value
                                                       MsgBox (LibR), vbOK
                                                       Exit Sub
                                                End If
   
   
   Set Sarray = REF.ListColumns("C224").DataBodyRange
                                                On Error Resume Next
                                                Set FRow = Sarray.Find(What:=TransB, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                If Not FRow Is Nothing Then
                                                        SCode = FRow.Row - 1
                                                        'Cahnge Value---------------------------------------------------------------------------
                                                        REF.ListColumns(C2).DataBodyRange(SCode, 1).value = CatB2
                                                Else
                                                   LibR = ThisWorkbook.Sheets("Library").Range("A59").value
                                                   MsgBox (LibR), vbOK
                                                   Exit Sub
                                                End If

   Row3 = Cat3TB.ListRows.Count
   Cat3TB.ListRows.Add
   Cat3TB.Resize Cat3TB.Range.Resize(Cat3TB.ListRows.Count + 1, Cat3TB.ListColumns.Count)
   Cat3TB.DataBodyRange(Row3 + 1, 1).value = TransB
   Set Sarray = Cat3TA.ListColumns(1).DataBodyRange
                                                 On Error Resume Next
                                                Set RowIndex = Sarray.Find(What:=TransB, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                If Not RowIndex Is Nothing Then
                                                        RowI = RowIndex.Row - 1
                                                        'Cahnge Value---------------------------------------------------------------------------
                                                        Cat3TA.ListRows(RowI).Delete
                                                End If
   Set Sarray3 = tb2.ListColumns(C3).DataBodyRange
   Set Sarray2 = tb2.ListColumns(C2).DataBodyRange
   Data3 = Sarray3.value
   Data2 = Sarray2.value
    
    ' Loop through the array to check and update values
    For i = 1 To UBound(Data3, 1)
        If Data3(i, 1) = TransB Then
            ' Update the value in column A
            Sarray2(i, 1).value = CatB2 ' Replace with your desired value
        End If
    Next i
    Set Cat3TA = ThisWorkbook.Sheets("WB_Data").ListObjects(CatA2)
    Set Cat3TB = ThisWorkbook.Sheets("WB_Data").ListObjects(CatB2)
    Me.Cat3A.RowSource = Cat3TA
    Me.Cat3B.RowSource = Cat3TB
    z = 1
   End If

Me.NFCB.Caption = ValNFCB
Me.NFCA.Caption = ValNFCA
Me.NFCA.ForeColor = &HFFFFFF
Me.NFCB.ForeColor = &HFFFFFF
End Sub
Private Sub TransfertAI_Click()
Dim reponse As VbMsgBoxResult
Dim Check As Long
Dim LibR, LibR1, LibR2 As String
Check = 0
Me.LabelCATI.Caption = ""

If Me.Cat3BI.ListIndex = -1 Then
   Exit Sub
End If


If Me.Cat2AI.value = "" Or Me.Cat2BI.value = "" Then
   Exit Sub
   Else
   Check = Cat3TBI.ListRows.Count
         If Check < 2 Then
         LibR = ThisWorkbook.Sheets("Library").Range("A65").value
         MsgBox (LibR), vbOK
         Exit Sub
         End If
   If A = 0 Then
   LibR1 = ThisWorkbook.Sheets("Library").Range("A66").value
   LibR2 = ThisWorkbook.Sheets("Library").Range("A67").value
   reponse = MsgBox((LibR1 & vbCrLf & vbCrLf & LibR2), vbOKCancel)
   If reponse = vbOK Then
   A = 1
   Else
   Exit Sub
   End If
   End If
   
   TransAI = Me.Cat3BI.value

                                                Set Sarray = Cat3TAI.ListColumns(1).DataBodyRange
                                                 On Error Resume Next
                                                Set RowIndex = Sarray.Find(What:=TransAI, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                If Not RowIndex Is Nothing Then
                                                       LibR = ThisWorkbook.Sheets("Library").Range("A57").value
                                                       MsgBox (LibR), vbOK
                                                       Exit Sub
                                                End If
                                                
   Set Sarray = REF.ListColumns("C224").DataBodyRange
                                                On Error Resume Next
                                                Set FRow = Sarray.Find(What:=TransAI, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                If Not FRow Is Nothing Then
                                                        SCode = FRow.Row - 1
                                                        'Cahnge Value---------------------------------------------------------------------------
                                                        REF.ListColumns(C2).DataBodyRange(SCode, 1).value = CatA2I
                                                Else
                                                   LibR = ThisWorkbook.Sheets("Library").Range("A59").value
                                                   MsgBox (LibR), vbOK
                                                   Exit Sub
                                                End If

   Row3 = Cat3TAI.ListRows.Count
   Cat3TAI.ListRows.Add
   Cat3TAI.Resize Cat3TAI.Range.Resize(Cat3TAI.ListRows.Count + 1, Cat3TAI.ListColumns.Count)
   Cat3TAI.DataBodyRange(Row3 + 1, 1).value = TransAI
   Set Sarray = Cat3TBI.ListColumns(1).DataBodyRange
                                                 On Error Resume Next
                                                Set RowIndex = Sarray.Find(What:=TransAI, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                If Not RowIndex Is Nothing Then
                                                        RowI = RowIndex.Row - 1
                                                        'Cahnge Value---------------------------------------------------------------------------
                                                        Cat3TBI.ListRows(RowI).Delete
                                                End If
   Set Sarray3 = tb2.ListColumns(C3).DataBodyRange
   Set Sarray2 = tb2.ListColumns(C2).DataBodyRange
   Data3 = Sarray3.value
   Data2 = Sarray2.value
    
    ' Loop through the array to check and update values
    For i = 1 To UBound(Data3, 1)
        If Data3(i, 1) = TransAI Then
            ' Update the value in column A
            Sarray2(i, 1).value = CatA2I ' Replace with your desired value
        End If
    Next i
    Set Cat3TAI = ThisWorkbook.Sheets("WB_Data").ListObjects(CatA2I)
    Set Cat3TBI = ThisWorkbook.Sheets("WB_Data").ListObjects(CatB2I)
    Me.Cat3AI.RowSource = Cat3TAI
    Me.Cat3BI.RowSource = Cat3TBI
    z = 1
   End If
End Sub
Private Sub TransfertBI_Click()
Dim reponse As VbMsgBoxResult
Dim Check As Long
Dim LibR, LibR1, LibR2 As String
Me.LabelCATI.Caption = ""
Check = 0

If Me.Cat3AI.ListIndex = -1 Then
   Exit Sub
End If

If Me.Cat2AI.value = "" Or Me.Cat2BI.value = "" Then
   Exit Sub
   Else
         Check = Cat3TAI.ListRows.Count
         If Check < 2 Then
         LibR = ThisWorkbook.Sheets("Library").Range("A65").value
         MsgBox (LibR), vbOK
         Exit Sub
         End If
   If A = 0 Then
   LibR1 = ThisWorkbook.Sheets("Library").Range("A66").value
   LibR2 = ThisWorkbook.Sheets("Library").Range("A67").value
   reponse = MsgBox((LibR1 & vbCrLf & vbCrLf & LibR2), vbOKCancel)
   If reponse = vbOK Then
   A = 1
   Else
   Exit Sub
   End If
   End If
   TransBI = Me.Cat3AI.value
                                                Set Sarray = Cat3TBI.ListColumns(1).DataBodyRange
                                                 On Error Resume Next
                                                Set RowIndex = Sarray.Find(What:=TransBI, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                If Not RowIndex Is Nothing Then
                                                       LibR = ThisWorkbook.Sheets("Library").Range("A57").value
                                                       MsgBox (LibR), vbOK
                                                       Exit Sub
                                                End If
   
   
   Set Sarray = REF.ListColumns("C224").DataBodyRange
                                                On Error Resume Next
                                                Set FRow = Sarray.Find(What:=TransBI, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                If Not FRow Is Nothing Then
                                                        SCode = FRow.Row - 1
                                                        'Cahnge Value---------------------------------------------------------------------------
                                                        REF.ListColumns(C2).DataBodyRange(SCode, 1).value = CatB2I
                                                Else
                                                   LibR = ThisWorkbook.Sheets("Library").Range("A59").value
                                                   MsgBox (LibR), vbOK
                                                   Exit Sub
                                                End If

   Row3 = Cat3TBI.ListRows.Count
   Cat3TBI.ListRows.Add
   Cat3TBI.Resize Cat3TBI.Range.Resize(Cat3TBI.ListRows.Count + 1, Cat3TBI.ListColumns.Count)
   Cat3TBI.DataBodyRange(Row3 + 1, 1).value = TransBI
   Set Sarray = Cat3TAI.ListColumns(1).DataBodyRange
                                                 On Error Resume Next
                                                Set RowIndex = Sarray.Find(What:=TransBI, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
                                                If Not RowIndex Is Nothing Then
                                                        RowI = RowIndex.Row - 1
                                                        'Cahnge Value---------------------------------------------------------------------------
                                                        Cat3TAI.ListRows(RowI).Delete
                                                End If
   Set Sarray3 = tb2.ListColumns(C3).DataBodyRange
   Set Sarray2 = tb2.ListColumns(C2).DataBodyRange
   Data3 = Sarray3.value
   Data2 = Sarray2.value
    
    ' Loop through the array to check and update values
    For i = 1 To UBound(Data3, 1)
        If Data3(i, 1) = TransBI Then
            ' Update the value in column A
            Sarray2(i, 1).value = CatB2I ' Replace with your desired value
        End If
    Next i
    Set Cat3TAI = ThisWorkbook.Sheets("WB_Data").ListObjects(CatA2I)
    Set Cat3TBI = ThisWorkbook.Sheets("WB_Data").ListObjects(CatB2I)
    Me.Cat3AI.RowSource = Cat3TAI
    Me.Cat3BI.RowSource = Cat3TBI
    z = 1
   End If

End Sub
Private Sub ChoiceC_Change()
Me.ChoiceC.BorderStyle = fmBorderStyleNone
Me.ChoiceC.ForeColor = vbWhite
LibR = ThisWorkbook.Sheets("Library").Range("A215").value
Me.Frame8.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A217").value
Me.Frame9.Caption = LibR
End Sub

Private Sub CmdSAVEC_Click()
Dim TbDataEP As ListObject
Dim tbSalaire As ListObject
Dim CELI As ListObject
Dim DTEP As ListObject
Dim chartObj As ChartObject

If Me.ChoiceC.value = "" Then
   Me.ChoiceC.BorderStyle = fmBorderStyleSingle
   Me.ChoiceC.BorderColor = vbRed
   Exit Sub
End If

             On Error Resume Next
             Set FRow = tbCCY.ListColumns("Currency").DataBodyRange.Find(What:=Me.ChoiceC.value, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
             If Not FRow Is Nothing Then
             CRY = Me.ChoiceC.value
             Else
             Me.ChoiceC.ForeColor = vbRed
             Exit Sub
             End If
tbCCY.ListColumns("Current").DataBodyRange(1, 1).value = CRY
Application.ScreenUpdating = False
ThisWorkbook.Sheets("ControlChart").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("ControlChart_Budget1").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("ControlChart_Budget2").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Sheet1").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value

Set TbDataEP = ThisWorkbook.Sheets("WB_Data").ListObjects("TbDataEP")
Set DTEP = ThisWorkbook.Sheets("ControlChart").ListObjects("DTEP")
Set tbSalaire = ThisWorkbook.Sheets("Sheet1").ListObjects("tbSalaire")
Set CELI = ThisWorkbook.Sheets("Sheet1").ListObjects("CELI")
TbDataEP.ListColumns("FIRE").DataBodyRange.NumberFormat = "# ### ### ### ##0 " & CRY
TbDataEP.ListColumns("Capital").DataBodyRange.NumberFormat = "# ### ### ### ##0 " & CRY
ThisWorkbook.Sheets("Analyse").Range("B30:C30").NumberFormat = "# ##0.00"
tbSalaire.ListColumns("Salaire annuel actuel").DataBodyRange.NumberFormat = "# ##0.00 " & CRY
CELI.ListColumns("First Year").DataBodyRange(4, 1).NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("ControlChart").Range("D21").NumberFormat = "# ##0.00"
ThisWorkbook.Sheets("ControlChart").Range("AS2:AZ13").NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("ControlChart").Range("AS18:AZ29").NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("ControlChart").Range("AS35:AZ46").NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("ControlChart").Range("AS35:AZ46").NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("ControlChart").Range("AW14").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart").Range("AZ14").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart").Range("AW30").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart").Range("AZ30").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart").Range("AW47").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart").Range("AZ47").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart").Range("BW1:BW7").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart").Range("CD1:CD7").NumberFormat = "(# ##0.00 " & CRY & ")"
ThisWorkbook.Sheets("ControlChart").Range("BW12:BW12").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart").Range("CG1").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart").Range("CM14").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart").Range("CM16").NumberFormat = "# ##0.00"
ThisWorkbook.Sheets("ControlChart").Range("CN3").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart").Range("CO4:CO5").NumberFormat = "# ##0.00 " & CRY
DTEP.ListColumns("Stratégie1").DataBodyRange.NumberFormat = "# ##0 " & CRY
DTEP.ListColumns("Stratégie2").DataBodyRange.NumberFormat = "# ##0 " & CRY
DTEP.ListColumns("Stratégie3").DataBodyRange.NumberFormat = "# ##0 " & CRY
DTEP.ListColumns("TTEP").DataBodyRange(1, 1).NumberFormat = "# ##0 " & CRY
DTEP.ListColumns("FI1").DataBodyRange.NumberFormat = "# ##0 " & CRY
DTEP.ListColumns("FI2").DataBodyRange.NumberFormat = "# ##0 " & CRY
DTEP.ListColumns("FI3").DataBodyRange.NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("ControlChart").Range("FG1").NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("ControlChart").Range("FK2").NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("ControlChart").Range("FL2").NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("ControlChart").Range("FN2").NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget1").Range("AT8:AU8").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget1").Range("AW8").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget1").Range("AT18:AU18").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget1").Range("AW18").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget1").Range("AT27:AU27").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget1").Range("AW27").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget1").Range("AT36:AU36").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget1").Range("AW36").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget1").Range("AT45:AU45").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget1").Range("AW45").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget1").Range("AT54:AU54").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget1").Range("AW54").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget1").Range("AT63:AU63").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget1").Range("AW63").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget1").Range("ES2:EU13").NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget1").Range("ES18:EU29").NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget1").Range("ES15:EU15").NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget1").Range("ES30:EU30").NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget2").Range("AT8:AU8").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget2").Range("AW8").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget2").Range("AT18:AU18").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget2").Range("AW18").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget2").Range("AT27:AU27").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget2").Range("AW27").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget2").Range("AT36:AU36").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget2").Range("AW36").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget2").Range("AT45:AU45").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget2").Range("AW45").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget2").Range("AT54:AU54").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget2").Range("AW54").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget2").Range("AT63:AU63").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget2").Range("AW63").NumberFormat = "# ##0.00 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget2").Range("ES2:EU13").NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget2").Range("ES18:EU29").NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget2").Range("ES15:EU15").NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("ControlChart_Budget2").Range("ES30:EU30").NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("Analyse").Range("B21:C21").NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("A_Budget1").Range("AB21:AC21").NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("A_Budget1").Range("AD23").NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("A_Budget2").Range("AB21:AC21").NumberFormat = "# ##0 " & CRY
ThisWorkbook.Sheets("A_Budget2").Range("AD23").NumberFormat = "# ##0 " & CRY
Set chartObj = ThisWorkbook.Sheets("DashBoard").ChartObjects("FlowChart")
chartObj.Chart.Axes(xlValue).TickLabels.NumberFormat = "# ##0 " & CRY
Set chartObj = ThisWorkbook.Sheets("DashBoard").ChartObjects("Flow Chart")
chartObj.Chart.Axes(xlValue).TickLabels.NumberFormat = "# ##0 " & CRY
Set chartObj = ThisWorkbook.Sheets("DashBoard").ChartObjects("Chart Strategique")
chartObj.Chart.Axes(xlValue).TickLabels.NumberFormat = "# ##0 " & CRY
Set chartObj = ThisWorkbook.Sheets("DashBoard").ChartObjects("Chart Prevision1")
chartObj.Chart.Axes(xlValue).TickLabels.NumberFormat = "# ##0 " & CRY
Set chartObj = ThisWorkbook.Sheets("DashBoard").ChartObjects("Chart Prevision2")
chartObj.Chart.Axes(xlValue).TickLabels.NumberFormat = "# ##0 " & CRY

ThisWorkbook.Sheets("ControlChart").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("ControlChart_Budget1").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("ControlChart_Budget2").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Sheet1").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
LibR = ThisWorkbook.Sheets("Library").Range("A216").value
Me.Frame8.Caption = LibR
ThisWorkbook.Sheets("DashBoard").Activate
Application.ScreenUpdating = True

End Sub
Private Sub UserForm_Initialize()
Dim LibR As String
Application.Calculation = xlCalculationAutomatic
LibR = ThisWorkbook.Sheets("Library").Range("A40").value
Me.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A41").value
Me.Comandquit.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A42").value
Me.AD1.Caption = LibR
Me.AD2.Caption = LibR
Me.AD1I.Caption = LibR
Me.AD2I.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A43").value
Me.ModifCat.Caption = LibR
Me.ModifCatB.Caption = LibR
Me.ModifCatI.Caption = LibR
Me.ModifCatBI.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A44").value
Me.Label1.Caption = LibR
Me.Label4.Caption = LibR
LibRD = ThisWorkbook.Sheets("Library").Range("A45").value
Me.LabelD.Caption = LibRD
LibRI = ThisWorkbook.Sheets("Library").Range("A46").value
Me.LabelI.Caption = LibRI
LibR = ThisWorkbook.Sheets("Library").Range("A47").value
Me.Label3.Caption = LibR
Me.Label8.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A49").value
Me.Valide.Caption = LibR
Me.ValideI.Caption = LibR
Me.VALIDLA.Caption = LibR
Me.VALIDERPATH.Caption = LibR
Me.CmdSAVEC.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A50").value
Me.LabelIn1.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A51").value
Me.LabelIn2.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A52").value
Me.LabelIn3.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A53").value
Me.Label11.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A54").value
Me.MultiPage1.Pages(0).Caption = LibR
Me.MultiPage1.Pages(2).Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A55").value
Me.MultiPage1.Pages(1).Caption = LibR
Me.MultiPage1.Pages(3).Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A56").value
Me.MultiPage1.Pages(4).Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A208").value
Me.Label12.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A215").value
Me.Frame8.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A217").value
Me.Frame9.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A213").value
Me.Label13.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A214").value
Me.Label14.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A187").value
Me.LNFF.Caption = LibR
Me.LNFFI.Caption = LibR


Set tbCCY = ThisWorkbook.Sheets("WB_Data").ListObjects("tbCCY")
CRY = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
Me.ChoiceC.value = CRY
Me.ChoiceL.value = tbCCY.ListColumns("ChoiceL").DataBodyRange(1, 1).value
Set Cat2 = ThisWorkbook.Sheets("WB_Data").ListObjects("Dépense")
Set Cat2I = ThisWorkbook.Sheets("WB_Data").ListObjects("Income")
Me.LPath.value = tbCCY.ListColumns("Path").DataBodyRange(2, 1).value
Me.LFolder.value = tbCCY.ListColumns("Path").DataBodyRange(1, 1).value

Me.Cat2A.RowSource = Cat2
Me.Cat2B.RowSource = Cat2
Me.GCAT2.RowSource = Cat2

Me.Cat2AI.RowSource = Cat2I
Me.Cat2BI.RowSource = Cat2I


Me.GCAT2M.value = ""
Me.GCAT2MI.value = ""
Me.Cat2A.value = ""
Me.Cat2B.value = ""
Me.Cat2AI.value = ""
Me.Cat2BI.value = ""
Me.LabelCAT.Caption = ""
Me.LabelCATI.Caption = ""
Me.GComplete.Caption = ""
Me.GCompleteI.Caption = ""
TransfertA.Caption = ChrW(&H2190)
TransfertB.Caption = ChrW(&H2192)
TransfertAI.Caption = ChrW(&H2190)
TransfertBI.Caption = ChrW(&H2192)

Set tbRefin = ThisWorkbook.Sheets("WB_Data").ListObjects("TbREFIncome")
Set REF = ThisWorkbook.Sheets("WB_Data").ListObjects("Refund")
Set tbh = ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead")
Set tbMonth = ThisWorkbook.Sheets("WB_Data").ListObjects("Month")
C2 = tbh.ListColumns("Header4").DataBodyRange(1, 1).value
C3 = tbh.ListColumns("Header5").DataBodyRange(1, 1).value
Set tb1 = ThisWorkbook.Sheets("Input").ListObjects("DataTable")
Set tb2 = ThisWorkbook.Sheets("Data Base").ListObjects("Master")
Set TbB1 = ThisWorkbook.Sheets("Budget1").ListObjects("DataBudget1")
Set TbEx1 = ThisWorkbook.Sheets("TamponBudget1").ListObjects("ExtrapBudget1")
Set TbTP1 = ThisWorkbook.Sheets("TamponBudget1").ListObjects("TamponBudget1")
Set TbEM1 = ThisWorkbook.Sheets("TamponBudget1").ListObjects("EMoisBudget1")
Set TbTPC1 = ThisWorkbook.Sheets("TamponBudget1").ListObjects("TbTPCBudget1")
Set TbB2 = ThisWorkbook.Sheets("Budget2").ListObjects("DataBudget2")
Set TbEx2 = ThisWorkbook.Sheets("TamponBudget2").ListObjects("ExtrapBudget2")
Set TbTP2 = ThisWorkbook.Sheets("TamponBudget2").ListObjects("TamponBudget2")
Set TbEM2 = ThisWorkbook.Sheets("TamponBudget2").ListObjects("EMoisBudget2")
Set TbTPC2 = ThisWorkbook.Sheets("TamponBudget2").ListObjects("TbTPCBudget2")
Set tbLmm = ThisWorkbook.Sheets("Library").ListObjects("tbLibrarymm")
Set tbLm = ThisWorkbook.Sheets("Library").ListObjects("tbLibrarym")
MPTASS = tbh.ListColumns("MP").DataBodyRange(1, 1).value
ValNFFC = REF.ListRows.Count
Me.NFFC.Caption = ValNFFC
Me.NFFCI.Caption = ValNFFC
Me.NFA.Visible = False
Me.NFB.Visible = False
Me.NFCA.Caption = ""
Me.NFCB.Caption = ""

z = 0
A = 0
End Sub

Private Sub Valide_Click()
Dim shp As Shape
Dim LibR As String

Err.Clear
ER = 0
Me.GComplete.Caption = ""
GCAT = ""
MODF = ""

MODF = Me.GCAT2.value

If MODF = "" Then
Me.GCAT2.BorderStyle = fmBorderStyleSingle
Me.GCAT2.BorderColor = vbRed
ER = 1
End If


GCAT = Me.GCAT2M.value
GCAT = Replace(Replace(GCAT, " ", "_"), Chr(160), "_")
If GCAT = "" Or GCAT = "Dépense" Then
Me.GCAT2M.BorderStyle = fmBorderStyleSingle
Me.GCAT2M.BorderColor = vbRed
Exit Sub
End If
Me.GCAT2M.value = GCAT
If ER = 1 Then
Exit Sub
End If

   Set Sarray = Cat2.ListColumns("C21").DataBodyRange
   Set FRow = Sarray.Find(What:=GCAT, LookIn:=xlValues, LookAt:=xlWhole)
               On Error Resume Next
               If Not FRow Is Nothing Then
                  LibR = ThisWorkbook.Sheets("Library").Range("A57").value
                  MsgBox (LibR), vbOK
                  Exit Sub
               End If


Set tbMOF = ThisWorkbook.Sheets("WB_Data").ListObjects(MODF)

    ' Change the name of the table
    On Error Resume Next
    tbMOF.Name = GCAT
    If Err.Number <> 0 Then
        Err.Clear
        Set tbMOF = ThisWorkbook.Sheets("WB_Data").ListObjects(MODF)
        LibR = ThisWorkbook.Sheets("Library").Range("A58").value
        MsgBox (LibR), vbExclamation
        Me.GCAT2M.ForeColor = vbRed
        Exit Sub
    End If
            
               
               Set FRow = Sarray.Find(What:=MODF, LookIn:=xlValues, LookAt:=xlWhole)
               On Error Resume Next
                                               If Not FRow Is Nothing Then
                                                        SCode = FRow.Row - 1
                                                        'Change Value---------------------------------------------------------------------------
                                                        Cat2.ListColumns("C21").DataBodyRange(SCode, 1).value = GCAT
                                                Else
                                                   LibR = ThisWorkbook.Sheets("Library").Range("A59").value
                                                   MsgBox (LibR), vbOK
                                                   Exit Sub
                                                End If
Me.MousePointer = fmMousePointerHourGlass
LibR = ThisWorkbook.Sheets("Library").Range("A60").value
Me.Caption = LibR

   Set Sarray = REF.ListColumns("Cat_2").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=GCAT, LookAt:=xlWhole
   Set Sarray = tb2.ListColumns(C2).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=GCAT, LookAt:=xlWhole
   Set Sarray = TbB1.ListColumns(C2).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=GCAT, LookAt:=xlWhole
   Set Sarray = TbEx1.ListColumns(C2).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=GCAT, LookAt:=xlWhole
   Set Sarray = TbB2.ListColumns(C2).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=GCAT, LookAt:=xlWhole
   Set Sarray = TbEx2.ListColumns(C2).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=GCAT, LookAt:=xlWhole
   MODF = "Btn " & MODF
   GCATB = "Btn " & GCAT

 For Each shp In ThisWorkbook.Sheets("WB_Data").Shapes
        If shp.Name = MODF Then
            ' Set the new name and caption
            shp.Name = GCATB
            shp.TextFrame.Characters.Text = GCAT
            Exit For
        End If
    Next shp
Me.GCAT2.value = ""
Me.GCAT2M.value = ""
Set Cat2 = ThisWorkbook.Sheets("WB_Data").ListObjects("Dépense")
Me.Cat2A.RowSource = Cat2
Me.Cat2B.RowSource = Cat2
Me.GCAT2.RowSource = Cat2
LibR = ThisWorkbook.Sheets("Library").Range("A61").value
Me.GComplete.Caption = LibR
z = 1
Me.MousePointer = fmMousePointerDefault
LibR = ThisWorkbook.Sheets("Library").Range("A40").value
Me.Caption = LibR
End Sub
Private Sub ValideI_Click()
Dim shp As Shape
Dim LibR As String
Err.Clear
ER = 0
Me.GCompleteI.Caption = ""
GCAT = ""
MODF = ""

MODF = Me.GCAT2I.value

If MODF = "" Then
Me.GCAT2I.BorderStyle = fmBorderStyleSingle
Me.GCAT2I.BorderColor = vbRed
ER = 1
End If

GCAT = Me.GCAT2MI.value
GCAT = Replace(Replace(GCAT, " ", "_"), Chr(160), "_")
GCAT = Replace(GCAT, " ", "_")
If GCAT = "" Or GCAT = "Income" Then
Me.GCAT2MI.BorderStyle = fmBorderStyleSingle
Me.GCAT2MI.BorderColor = vbRed
Exit Sub
End If
Me.GCAT2MI.value = GCAT

Select Case PINC
       Case Is = 1
       GCAT = "i1_" & Me.GCAT2MI.value
       Case Is = 2
       GCAT = "i2_" & Me.GCAT2MI.value
       Case Is = 3
       GCAT = "i3_" & Me.GCAT2MI.value
       Case Is = ""
       Me.GCAT2I.BorderStyle = fmBorderStyleSingle
       Me.GCAT2I.BorderColor = vbRed
       ER = 1
       Case Is = 0
       Me.GCAT2I.BorderStyle = fmBorderStyleSingle
       Me.GCAT2I.BorderColor = vbRed
       ER = 1
End Select





If ER = 1 Then
Exit Sub
End If


   Set Sarray = Cat2I.ListColumns("C22").DataBodyRange
   Set FRow = Sarray.Find(What:=GCAT, LookIn:=xlValues, LookAt:=xlWhole)
               On Error Resume Next
               If Not FRow Is Nothing Then
                  LibR = ThisWorkbook.Sheets("Library").Range("A57").value
                  MsgBox (LibR), vbOK
                  Exit Sub
               End If

MODF = Cat2I.ListColumns("C22").DataBodyRange(PINC, 1).value

Set tbMOF = ThisWorkbook.Sheets("WB_Data").ListObjects(MODF)

    ' Change the name of the table
    On Error Resume Next
    tbMOF.Name = GCAT
    If Err.Number <> 0 Then
        Set tbMOF = ThisWorkbook.Sheets("WB_Data").ListObjects(MODF)
        Me.GCAT2MI.ForeColor = vbRed
        LibR = ThisWorkbook.Sheets("Library").Range("A58").value
        MsgBox (LibR), vbExclamation
        Exit Sub
    End If



               Set FRow = Sarray.Find(What:=MODF, LookIn:=xlValues, LookAt:=xlWhole)
               On Error Resume Next
                                               If Not FRow Is Nothing Then
                                                        SCode = FRow.Row - 1
                                                        'Change Value---------------------------------------------------------------------------
                                                        Cat2I.ListColumns("C22").DataBodyRange(SCode, 1).value = GCAT
                                                Else
                                                   LibR = ThisWorkbook.Sheets("Library").Range("A59").value
                                                   MsgBox (LibR), vbOK
                                                   Exit Sub
                                                End If

Application.ScreenUpdating = False
Me.MousePointer = fmMousePointerHourGlass
LibR = ThisWorkbook.Sheets("Library").Range("A60").value
Me.Caption = LibR

   Set Sarray = REF.ListColumns("Cat_2").DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=GCAT, LookAt:=xlWhole
   Set Sarray = tb2.ListColumns(C2).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=GCAT, LookAt:=xlWhole
   Set Sarray = TbB1.ListColumns(C2).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=GCAT, LookAt:=xlWhole
   Set Sarray = TbEx1.ListColumns(C2).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=GCAT, LookAt:=xlWhole
   Set Sarray = TbB2.ListColumns(C2).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=GCAT, LookAt:=xlWhole
   Set Sarray = TbEx2.ListColumns(C2).DataBodyRange
   Sarray.Replace What:=MODF, Replacement:=GCAT, LookAt:=xlWhole
   MODF = "Btn " & MODF
   GCATB = "Btn " & GCAT

 For Each shp In ThisWorkbook.Sheets("WB_Data").Shapes
        If shp.Name = MODF Then
            ' Set the new name and caption
            shp.Name = GCATB
            shp.TextFrame.Characters.Text = GCAT
            Exit For
        End If
    Next shp
Me.GCAT2I.value = ""
Me.GCAT2MI.value = ""
Set Cat2I = ThisWorkbook.Sheets("WB_Data").ListObjects("Income")
Me.Cat2AI.RowSource = Cat2I
Me.Cat2BI.RowSource = Cat2I
LibR = ThisWorkbook.Sheets("Library").Range("A61").value
Me.GCompleteI.Caption = LibR
z = 1
Me.LabelIn1.ForeColor = &H808080
Me.LabelIn2.ForeColor = &H808080
Me.LabelIn3.ForeColor = &H808080


Me.MousePointer = fmMousePointerDefault
LibR = ThisWorkbook.Sheets("Library").Range("A40").value
Me.Caption = LibR
Application.ScreenUpdating = True
End Sub

Private Sub VALIDERPATH_Click()
ER = 0
If Me.LPath.value = "" Then
   Me.LPath.BorderStyle = fmBorderStyleSingle
   Me.LPath.BorderColor = vbRed
   ER = 1
End If
If Me.LFolder.value = "" Then
   Me.LFolder.BorderStyle = fmBorderStyleSingle
   Me.LFolder.BorderColor = vbRed
   ER = 1
End If
If ER = 1 Then
Exit Sub
End If
tbCCY.ListColumns("Path").DataBodyRange(2, 1).value = Me.LPath.value
tbCCY.ListColumns("Path").DataBodyRange(1, 1).value = Me.LFolder.value
tbCCY.ListColumns("Module").DataBodyRange(1, 1).value = Me.LFolder.value
LibR = ThisWorkbook.Sheets("Library").Range("A218").value
Me.Frame9.Caption = LibR
End Sub

Private Sub VALIDLA_Click()
Dim shp As Shape
Dim LibR, chineseText, LangVal, langCode, Separate As String
Dim Origin, NewV, Stringsl, pfString, LastWth, LastInc, MO1 As String
Dim NewSML  As Variant
Dim SMArray() As Variant
Dim SMMMArray() As Variant
Dim Sarraymmmm As Range
Dim Sarraym3 As Range
Dim ptArray() As Variant
Dim shtArray() As Variant
Dim SarrayIN, SarrayB1, SarrayB2, SarrayMASTERm, FRow As Range
Dim pt As PivotTable
Dim pf As PivotField
Dim pi As PivotItem
Dim slc As Slicer
Dim slicerItem As slicerItem
Dim slicerField As slicerCache
Dim v, k, YRow, CCP As Integer
ReDim SMArray(1 To 12)
ReDim SMMMArray(1 To 12)

y = 18
For i = 1 To 12
    SMArray(i) = tbMonth.ListColumns("Mois").DataBodyRange(i, 1).value
    SMMMArray(i) = ThisWorkbook.Sheets("Library").Cells(y, "R").value
    y = y + 1
Next i

If Me.ChoiceL.value = "" Then
   Me.ChoiceL.BorderStyle = fmBorderStyleSingle
   Me.ChoiceL.BorderColor = vbRed
   Exit Sub
End If
Origin = tbCCY.ListColumns("ChoiceL").DataBodyRange(1, 1).value
If Origin = Me.ChoiceL.value Then
   Exit Sub
End If

             On Error Resume Next
             Set FRow = tbCCY.ListColumns("Language").DataBodyRange.Find(What:=Me.ChoiceL.value, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious, SearchOrder:=xlByRows)
             If Not FRow Is Nothing Then
             LangVal = Me.ChoiceL.value
             tbCCY.ListColumns("ChoiceL").DataBodyRange(1, 1).value = LangVal
             Else
             Me.ChoiceL.ForeColor = vbRed
             Exit Sub
             End If

CCP = 0
Me.MousePointer = fmMousePointerHourGlass
LibR = ThisWorkbook.Sheets("Library").Range("A60").value
Me.Caption = LibR
Application.ScreenUpdating = False
ThisWorkbook.Sheets("DashBoard").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value


LibR = ThisWorkbook.Sheets("Library").Range("A177").value
Set shp = ThisWorkbook.Sheets("DashBoard").Shapes("Rectangle: Rounded Corners 5410")
shp.TextFrame2.TextRange.Text = LibR
Set shp = ThisWorkbook.Sheets("DashBoard").Shapes("Rectangle: Rounded Corners 5409")
shp.TextFrame2.TextRange.Text = LibR
Set shp = ThisWorkbook.Sheets("DashBoard").Shapes("Rectangle: Rounded Corners 5398")
shp.TextFrame2.TextRange.Text = LibR
Set shp = ThisWorkbook.Sheets("DashBoard").Shapes("Rectangle: Rounded Corners 5395")
shp.TextFrame2.TextRange.Text = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A212").value
Set shp = ThisWorkbook.Sheets("DashBoard").Shapes("Rectangle: Rounded Corners 5869")
shp.TextFrame2.TextRange.Text = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A179").value
Set shp = ThisWorkbook.Sheets("DashBoard").Shapes("Rectangle: Rounded Corners 5412")
shp.TextFrame2.TextRange.Text = LibR
Set shp = ThisWorkbook.Sheets("DashBoard").Shapes("Rectangle: Rounded Corners 5834")
shp.TextFrame2.TextRange.Text = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A228").value
Set shp = ThisWorkbook.Sheets("DashBoard").Shapes("Rectangle: Rounded Corners 5884")
shp.TextFrame2.TextRange.Text = LibR

LibR = ThisWorkbook.Sheets("Library").Range("A178").value
Set shp = ThisWorkbook.Sheets("DashBoard").Shapes("Rectangle: Rounded Corners 5413")
shp.TextFrame2.TextRange.Text = LibR

LibR = ThisWorkbook.Sheets("Library").Range("A180").value
Set shp = ThisWorkbook.Sheets("DashBoard").Shapes("Rectangle: Rounded Corners 5415")
shp.TextFrame2.TextRange.Text = LibR
Set shp = ThisWorkbook.Sheets("DashBoard").Shapes("Rectangle: Rounded Corners 5387")
shp.TextFrame2.TextRange.Text = LibR
Set shp = ThisWorkbook.Sheets("DashBoard").Shapes("Rectangle: Rounded Corners 5886")
shp.TextFrame2.TextRange.Text = LibR

LibR = ThisWorkbook.Sheets("Library").Range("A181").value
Set shp = ThisWorkbook.Sheets("DashBoard").Shapes("Rectangle: Rounded Corners 5486")
shp.TextFrame2.TextRange.Text = LibR


LibR = ThisWorkbook.Sheets("Library").Range("A182").value
Set shp = ThisWorkbook.Sheets("DashBoard").Shapes("Shape Strat1")
shp.TextFrame2.TextRange.Text = LibR
Set shp = ThisWorkbook.Sheets("DashBoard").Shapes("ShapeStrat11")
shp.TextFrame2.TextRange.Text = LibR

LibR = ThisWorkbook.Sheets("Library").Range("A183").value
Set shp = ThisWorkbook.Sheets("DashBoard").Shapes("ShapeStrat22")
shp.TextFrame2.TextRange.Text = LibR
Set shp = ThisWorkbook.Sheets("DashBoard").Shapes("Shape Strat2")
shp.TextFrame2.TextRange.Text = LibR

LibR = ThisWorkbook.Sheets("Library").Range("A184").value
Set shp = ThisWorkbook.Sheets("DashBoard").Shapes("ShapeStrat33")
shp.TextFrame2.TextRange.Text = LibR
Set shp = ThisWorkbook.Sheets("DashBoard").Shapes("Shape Strat3")
shp.TextFrame2.TextRange.Text = LibR

LibR = ThisWorkbook.Sheets("Library").Range("A41").value
Me.Comandquit.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A42").value
Me.AD1.Caption = LibR
Me.AD2.Caption = LibR
Me.AD1I.Caption = LibR
Me.AD2I.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A43").value
Me.ModifCat.Caption = LibR
Me.ModifCatB.Caption = LibR
Me.ModifCatI.Caption = LibR
Me.ModifCatBI.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A44").value
Me.Label1.Caption = LibR
Me.Label4.Caption = LibR
LibRD = ThisWorkbook.Sheets("Library").Range("A45").value
Me.LabelD.Caption = LibRD
LibRI = ThisWorkbook.Sheets("Library").Range("A46").value
Me.LabelI.Caption = LibRI
LibR = ThisWorkbook.Sheets("Library").Range("A47").value
Me.Label3.Caption = LibR
Me.Label8.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A49").value
Me.Valide.Caption = LibR
Me.ValideI.Caption = LibR
Me.VALIDLA.Caption = LibR
Me.VALIDERPATH.Caption = LibR
Me.CmdSAVEC.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A50").value
Me.LabelIn1.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A51").value
Me.LabelIn2.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A52").value
Me.LabelIn3.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A53").value
Me.Label11.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A54").value
Me.MultiPage1.Pages(0).Caption = LibR
Me.MultiPage1.Pages(2).Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A55").value
Me.MultiPage1.Pages(1).Caption = LibR
Me.MultiPage1.Pages(3).Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A56").value
Me.MultiPage1.Pages(4).Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A208").value
Me.Label12.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A217").value
Me.Frame9.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A213").value
Me.Label13.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A214").value
Me.Label14.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A191").value
LastWth = tbRefin.ListColumns("SuperCat").DataBodyRange(1, 1).value
tbRefin.ListColumns("SuperCat").DataBodyRange(1, 1).value = LibR
LastInc = tbRefin.ListColumns("SuperCat").DataBodyRange(2, 1).value
LibR = ThisWorkbook.Sheets("Library").Range("A187").value
Me.LNFF.Caption = LibR
Me.LNFFI.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A211").value
tbRefin.ListColumns("SuperCat").DataBodyRange(2, 1).value = LibR


chineseText = ChrW(&H4E2D) & ChrW(&H6587)
Select Case LangVal
        Case "Français", "fr"
            langCode = "[$-40C]mmmm" ' French (janv., févr., mars...)
            Set Sarraymmmm = tbLmm.ListColumns(LangVal).DataBodyRange
            Set Sarraym3 = tbLm.ListColumns(LangVal).DataBodyRange
        Case "English", "en"
            langCode = "[$-409]mmmm" ' English (Jan, Feb, Mar...)
            Set Sarraymmmm = tbLmm.ListColumns(LangVal).DataBodyRange
            Set Sarraym3 = tbLm.ListColumns(LangVal).DataBodyRange
        Case "Español", "es"
            langCode = "[$-C0A]mmmm" ' Spanish (ene., feb., mar...)
            Set Sarraymmmm = tbLmm.ListColumns(LangVal).DataBodyRange
            Set Sarraym3 = tbLm.ListColumns(LangVal).DataBodyRange
        Case chineseText, "zh"
            langCode = "[$-804]mmmm" ' Simplified Chinese (1?, 2?, 3?...)
            Set Sarraymmmm = tbLmm.ListColumns(LangVal).DataBodyRange
            Set Sarraym3 = tbLm.ListColumns(LangVal).DataBodyRange
            CCP = 1
        Case "Deutsch", "zh"
            langCode = "[$-407]mmmm" ' Simplified Chinese (1?, 2?, 3?...)
            Set Sarraymmmm = tbLmm.ListColumns(LangVal).DataBodyRange
            Set Sarraym3 = tbLm.ListColumns(LangVal).DataBodyRange
        Case Else
            MsgBox "Invalid language! Please enter French, English, Spanish, German or Chinese.", vbExclamation
            Exit Sub
    End Select

shtArray = Array("A_Budget1", "A_Budget2")
For k = LBound(shtArray) To UBound(shtArray)
Mysheet = shtArray(z)
Select Case k
       Case Is = 0
       ptArray = Array("Budget1_CashFlow", "Budget1_Cat3-1", "Budget1_Cat3-2")
       pfString = "[Budget1Data].[Mois (Month)].[Mois (Month)]"
       Separate = "[Budget1Data].[Mois (Month)].&["
       Case Is = 1

       ptArray = Array("Budget2_CashFlow", "Budget2_Cat3-1", "Budget2_Cat3-2")
       pfString = "[DataBudget2].[Mois (Month)].[Mois (Month)]"
       Separate = "[DataBudget2].[Mois (Month)].&["
End Select


For v = LBound(ptArray) To UBound(ptArray)

Stringpt = ptArray(v)
Set pt = ThisWorkbook.Sheets(Mysheet).PivotTables(Stringpt)
Set pf = pt.PivotFields(pfString)

y = 1
 For Each pi In pf.PivotItems
        For i = y To 12
        Origin = Separate & tbLm.ListColumns("Français").DataBodyRange(i, 1).value & "]"
        If pi.value = Origin Then
            NewV = tbLm.ListColumns(LangVal).DataBodyRange(i, 1).value
            pi.value = NewV ' Replace French month name with English
            Exit For
        End If
        Next i
    Next pi
Next v


Next k

tb2.ListColumns("Mois").DataBodyRange.NumberFormat = langCode
TbB1.ListColumns("Mois").DataBodyRange.NumberFormat = langCode
TbB2.ListColumns("Mois").DataBodyRange.NumberFormat = langCode
TbEx1.ListColumns("Mois").DataBodyRange.NumberFormat = langCode
TbEx2.ListColumns("Mois").DataBodyRange.NumberFormat = langCode
Set SarrayIN = tb1.ListColumns("Mois").DataBodyRange
Set SarrayB1 = TbEM1.ListColumns("Mois").DataBodyRange
Set SarrayB2 = TbEM2.ListColumns("Mois").DataBodyRange
Set SarrayMASTERm = tb2.ListColumns("Mois (Name)").DataBodyRange


For i = 1 To 12
   SarrayIN.Replace What:=SMArray(i), Replacement:=Sarraymmmm(i, 1), LookAt:=xlWhole
   SarrayB1.Replace What:=SMArray(i), Replacement:=Sarraymmmm(i, 1), LookAt:=xlWhole
   SarrayB2.Replace What:=SMArray(i), Replacement:=Sarraymmmm(i, 1), LookAt:=xlWhole
   SarrayMASTERm.Replace What:=SMMMArray(i), Replacement:=Sarraym3(i, 1), LookAt:=xlWhole
Next i
Set SarrayIN = tb1.ListColumns("Cat_1").DataBodyRange
    SarrayIN.Replace What:=LastWth, Replacement:=tbRefin.ListColumns("SuperCat").DataBodyRange(1, 1).value, LookAt:=xlWhole
    SarrayIN.Replace What:=LastInc, Replacement:=tbRefin.ListColumns("SuperCat").DataBodyRange(2, 1).value, LookAt:=xlWhole


ThisWorkbook.Sheets("DashBoard").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
Me.MousePointer = fmMousePointerDefault
LibR = ThisWorkbook.Sheets("Library").Range("A40").value
Me.Caption = LibR
DisableEvents = True
Me.LPath.value = tbCCY.ListColumns("Path").DataBodyRange(2, 1).value
Me.LFolder.value = tbCCY.ListColumns("Path").DataBodyRange(1, 1).value
DisableEvents = False
LibR = ThisWorkbook.Sheets("Library").Range("A216").value
Me.Frame8.Caption = LibR
z = 1
ThisWorkbook.Sheets("DashBoard").Activate
Application.ScreenUpdating = True
End Sub
