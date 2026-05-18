VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputData 
   Caption         =   "Enter Data"
   ClientHeight    =   7170
   ClientLeft      =   112
   ClientTop       =   448
   ClientWidth     =   12728
   OleObjectBlob   =   "InputData.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "InputData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public helperRange, Autoarray As Range
Public Cat2 As ListObject
Public Cat3 As ListObject
Public tb As ListObject
Public tbh As ListObject
Public tbCCY As ListObject
Public tbRefin As ListObject
Public Rows As Long
Public Cash, V6 As Double
Public Start, EndE, arrIndex, selectedRow, i, ARow, AutoStart  As Long
Public SelectedValue, V2, V3, V4, V5, CRY, CodeM, Expense, SIncome, Math, Savings, localDec, altDec, MyMontant, vbaSep As String
Public A, V1, ER, lastIndex, AutoM, AUTO1, Tag, NewOne As Integer
Public ValM, VMontant As Variant
Dim DisableEvents As Boolean

Private Sub Add_Click()
Dim LibR As String
Dim Cash As Double
ER = 0
If Me.Montant.value = "" Then
   Me.Montant.BorderStyle = fmBorderStyleSingle
   Me.Montant.BorderColor = vbRed
   ER = 1
Else
    VMontant = Replace(Me.Montant.value, localDec, altDec)
    If vbaSep = "." Then
    VMontant = Replace(VMontant, altDec, Application.International(xlDecimalSeparator))
    End If
    VMontant = Trim(VMontant)
    If IsNumeric(VMontant) Then
        Cash = CDbl(VMontant)
    Else
    Me.Montant.ForeColor = vbRed
    ER = 1
    End If
End If

If Me.Year.value = "" Then
    Me.Year.BorderStyle = fmBorderStyleSingle
    Me.Year.BorderColor = vbRed
    ER = 1
End If

If Me.Cat_2.ListIndex = -1 Then
    Me.LCAT2.BorderStyle = fmBorderStyleSingle
    Me.LCAT2.BorderColor = vbRed
    ER = 1
End If

If Me.Cat_3.ListIndex = -1 Then
    Me.LCAT3.BorderStyle = fmBorderStyleSingle
    Me.LCAT3.BorderColor = vbRed
    ER = 1
End If

If ER = 1 Then
   Exit Sub
End If

If tbh.ListColumns("Max1").DataBodyRange(1, 1).value = 1 And tb.ListColumns("Année").DataBodyRange(1, 1).value = "Exemple" Then
   Rows = 1
   tb.ListColumns("Année").DataBodyRange(1, 1).Font.color = vbBlack
Else
   Rows = tbh.ListColumns("Max1").DataBodyRange(1, 1).value + 1
End If

If Me.Year.value <> "" And Me.Mois.ListIndex >= 0 And Me.Cat_1.ListIndex >= 0 And Me.Cat_2.ListIndex >= 0 And Me.Cat_3.ListIndex >= 0 Then

   tb.ListColumns("Année").DataBodyRange(Rows, 1).value = Me.Year.value
   tb.ListColumns("Mois").DataBodyRange(Rows, 1).value = Me.Mois.value
   tb.ListColumns("Cat_1").DataBodyRange(Rows, 1).value = Me.Cat_1.value
   tb.ListColumns("Cat_2").DataBodyRange(Rows, 1).value = Me.Cat_2.value
   tb.ListColumns("Cat_3").DataBodyRange(Rows, 1).value = Me.Cat_3.value
   Select Case Me.Cat_1.value
          Case Is = Expense
               If Cash > 0 Then
                  Cash = Cash * -1
               End If
          Case Is = SIncome
              If Cash < 0 Then
                  Cash = Cash * -1
               End If
   End Select
   tb.ListColumns("Montant").DataBodyRange(Rows, 1).value = Cash
   Me.Last_Cash.Caption = Cash
   Me.Montant.value = ""
   Me.PPA.value = ""
   
   
    If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets("Input").Range("Z:Z")) = 0 Then
    selectedRow = 1 ' If column Z is empty, start at row 1
    Else
    selectedRow = ThisWorkbook.Sheets("Input").Cells(ThisWorkbook.Sheets("Input").Rows.Count, "Z").End(xlUp).Row
    selectedRow = selectedRow + 1
    End If
   
   SelectedValue = Me.Year.value & "-" & Left(Me.Mois.value, 3) & "-" & Me.Cat_3.value & "-(" & Cash & CRY & ")"
   helperRange.Cells(selectedRow, 1).value = SelectedValue
   Set helperRange = ThisWorkbook.Sheets("Input").Range("Z1:Z" & ThisWorkbook.Sheets("Input").Cells(ThisWorkbook.Sheets("Input").Rows.Count, "Z").End(xlUp).Row)
   Me.PPA.RowSource = "'" & helperRange.Worksheet.Name & "'!" & helperRange.Address
   Me.PPA.TopIndex = Me.PPA.ListCount - 1
   
   
   
   
Else
LibR = ThisWorkbook.Sheets("Library").Range("A104").value
MsgBox (LibR), vbOKOnly
End If

If Tag = 1 And AUTO1 = 1 Then
        ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "F").value = 1
        AutoStart = AutoStart + 1 ' Default to the first row if no blank cells
        If AutoStart > ARow Then
        Me.LAuto.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 1, "B").value
        Me.LAuto.Font.Strikethrough = True
        Me.LAuto.ForeColor = &H808080
        LibR = ThisWorkbook.Sheets("Library").Range("A219").value
        Me.FrameAuto.Caption = LibR
        Me.FrameAuto.ForeColor = &HC0C0C0
        Me.Label6.ForeColor = &HC0C0C0
        Me.Montant.value = ""
        Me.CommandNext.Enabled = False
        AUTO1 = 0
        Exit Sub
        End If
        Me.CommandButton1.Enabled = True
        If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "F").value = 1 Then
        Me.LAuto.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "B").value
        Me.LAuto.Font.Strikethrough = 0
        Me.LAuto.ForeColor = &H808080
        LibR = ThisWorkbook.Sheets("Library").Range("A219").value
        Me.FrameAuto.Caption = LibR
        Me.FrameAuto.ForeColor = &HC0C0C0
        Me.Montant.value = ""
        AUTO1 = 0
        Else
        Me.LAuto.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "B").value
        Me.LAuto.Font.Strikethrough = 0
        Me.LAuto.ForeColor = &HFFFFC0
        LibR = ThisWorkbook.Sheets("Library").Range("A212").value
        Me.FrameAuto.Caption = LibR
        AUTO1 = 1
        End If
        If AutoStart + 1 <= ARow Then
           If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 1, "F").value = 1 Then
           Me.LN1.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 1, "B").value
           Me.LN1.Font.Strikethrough = True
           Me.LN1.ForeColor = &H808080
           Else
           Me.LN1.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 1, "B").value
           Me.LN1.Font.Strikethrough = 0
           Me.LN1.ForeColor = &HFFFFFF
           End If
        Else
        Me.LN1.Caption = ""
        End If
        If AutoStart + 2 <= ARow Then
           If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 2, "F").value = 1 Then
           Me.LN2.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 2, "B").value
           Me.LN2.Font.Strikethrough = True
           Me.LN2.ForeColor = &H808080
           Else
           Me.LN2.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 2, "B").value
           Me.LN2.Font.Strikethrough = 0
           Me.LN2.ForeColor = &HFFFFFF
           End If
        Else
        Me.LN2.Caption = ""
        End If
        If AutoStart - 1 >= 1 Then
           If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 1, "F").value = 1 Then
           Me.LL1.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 1, "B").value
           Me.LL1.Font.Strikethrough = True
           Me.LL1.ForeColor = &H808080
           Else
           Me.LL1.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 1, "B").value
           Me.LL1.Font.Strikethrough = 0
           Me.LL1.ForeColor = &HFFFFFF
           End If
        Else
        Me.LL1.Caption = ""
        End If
        If AutoStart - 2 >= 1 Then
           If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 2, "F").value = 1 Then
           Me.LL2.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 2, "B").value
           Me.LL2.Font.Strikethrough = True
           Me.LL2.ForeColor = &H808080
           Else
           Me.LL2.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 2, "B").value
           Me.LL2.Font.Strikethrough = 0
           Me.LL2.ForeColor = &HFFFFFF
           End If
        Else
        Me.LL2.Caption = ""
        End If

If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "C").value <> "" And AUTO1 = 1 Then
   DisableEvents = True
   Me.Montant.value = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "C").value
   DisableEvents = False
   Me.Label6.ForeColor = &HFFFFC0
   Me.FrameAuto.ForeColor = &HFFFFC0
   If Me.Cat_1.value = SIncome Then
   Me.Cat_1.value = Expense
   End If
End If
If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "D").value <> "" Then
   DisableEvents = True
   Me.Montant.value = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "D").value
   DisableEvents = False
   Me.Label6.ForeColor = &HFFFFC0
   Me.FrameAuto.ForeColor = &HFFFFC0
   If Me.Cat_1.value = Expense Then
   Me.Cat_1.value = SIncome
   End If
End If
AutoM = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "E").value
FirstMonth = ThisWorkbook.Sheets("Library").Cells(1 + AutoM, "R").value
On Error Resume Next
Me.Mois.value = FirstMonth
On Error GoTo 0
End If
End Sub

Private Sub Additional_Click()
AddAn.Show
Me.Year.RowSource = ThisWorkbook.Sheets("WB_Data").ListObjects("Year")
If AddAn.An.value <> "" Then
Me.Year = AddAn.An.value
Me.Mois.ListIndex = 0
End If
On Error Resume Next
Unload (AddAn)
End Sub

Private Sub Cancel_Click()

Rows = tbh.ListColumns("Max1").DataBodyRange(1, 1).value
If tb.ListColumns("Année").DataBodyRange(1, 1).value <> "Exemple" Then
   
Set rng = tb.ListRows(Rows).Range
rng.Delete
Me.PPA.value = ""
If Rows = 1 Then
Me.Last_Cash.Caption = "---"
   tb.ListColumns("Année").Range(2, 1).value = "Exemple"
Else

Rows = Rows - 1
Me.Last_Cash.Caption = tb.ListColumns("Montant").DataBodyRange(Rows, 1).value
End If
   If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets("Input").Range("Z:Z")) = 0 Then
    selectedRow = 1 ' If column Z is empty, start at row 1
    Else
    selectedRow = ThisWorkbook.Sheets("Input").Cells(ThisWorkbook.Sheets("Input").Rows.Count, "Z").End(xlUp).Row
    End If
   helperRange.Cells(selectedRow, 1).value = ""
   Set helperRange = ThisWorkbook.Sheets("Input").Range("Z1:Z" & ThisWorkbook.Sheets("Input").Cells(ThisWorkbook.Sheets("Input").Rows.Count, "Z").End(xlUp).Row)
   Me.PPA.RowSource = "'" & helperRange.Worksheet.Name & "'!" & helperRange.Address
   Me.PPA.TopIndex = Me.PPA.ListCount - 1
End If

End Sub

Private Sub Cat_1_Change()
If Me.Cat_1.ListIndex >= 0 Then
Me.Cat_1.BorderStyle = fmBorderStyleNone
Me.PPA.value = ""
Me.Cat_3.value = ""
Me.Cat_2.value = ""
CatB1 = Me.Cat_1.value
LibR = ThisWorkbook.Sheets("Library").Range("A245").value
Me.Add.Caption = LibR
If CatB1 = "" Then
Me.Test.Caption = ""
ElseIf CatB1 = Expense Then
CatB1 = "Dépense"
Me.Test.Enabled = True
LibR = ThisWorkbook.Sheets("Library").Range("A246").value
Me.Test.Caption = LibR
ElseIf CatB1 = SIncome Then
CatB1 = "Income"
LibR = ThisWorkbook.Sheets("Library").Range("A247").value
Me.Test.Caption = LibR
Me.Test.Enabled = False
End If
Set Cat2 = ThisWorkbook.Sheets("WB_Data").ListObjects(CatB1)
Me.Cat_2.RowSource = Cat2
End If
End Sub

Private Sub Cat_2_Change()

If Me.Cat_2.ListIndex >= 0 Then
Me.LCAT2.BorderStyle = fmBorderStyleNone
Me.PPA.value = ""
Me.Cat_3.value = ""
CatB2 = Me.Cat_2.value
If CatB2 = Savings Then
Me.Test.Enabled = True
LibR = ThisWorkbook.Sheets("Library").Range("A252").value
Me.Add.Caption = LibR
Else
  If Me.Cat_1.value = SIncome Then
     Me.Test.Enabled = False
     LibR = ThisWorkbook.Sheets("Library").Range("A245").value
     Me.Add.Caption = LibR
   End If
End If
Set Cat3 = ThisWorkbook.Sheets("WB_Data").ListObjects(CatB2)
Me.Cat_3.RowSource = Cat3
End If
End Sub

Private Sub Cat_3_Change()
If Me.Cat_3.ListIndex >= 0 Then
Me.LCAT3.BorderStyle = fmBorderStyleNone
Me.PPA.value = ""
End If
End Sub

Private Sub CommandButton1_Click()
AUTO1 = 1
If Tag = 1 Then
   
        AutoStart = AutoStart - 1 ' Default to the first row if no blank cells
        If AutoStart < 1 Then
        LibR = ThisWorkbook.Sheets("Library").Range("A219").value
        Me.FrameAuto.Caption = LibR
        Me.Label6.ForeColor = &HC0C0C0
        Me.FrameAuto.ForeColor = &HC0C0C0
        AUTO1 = 0
        AutoStart = 1
        Me.Montant.value = ""
        Me.CommandButton1.Enabled = False
        Exit Sub
        End If
        Me.CommandNext.Enabled = True
        If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "F").value = 1 Then
        Me.LAuto.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "B").value
        Me.LAuto.Font.Strikethrough = True
        Me.LAuto.ForeColor = &H808080
        Me.Label6.ForeColor = &HC0C0C0
        Me.FrameAuto.ForeColor = &HC0C0C0
        LibR = ThisWorkbook.Sheets("Library").Range("A219").value
        Me.FrameAuto.Caption = LibR
        Me.Montant.value = ""
        AUTO1 = 0
        Else
        Me.LAuto.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "B").value
        Me.LAuto.Font.Strikethrough = 0
        Me.LAuto.ForeColor = &HFFFFC0
        LibR = ThisWorkbook.Sheets("Library").Range("A212").value
        Me.FrameAuto.Caption = LibR
        AUTO1 = 1
        End If
        If AutoStart + 1 <= ARow Then
           If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 1, "F").value = 1 Then
           Me.LN1.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 1, "B").value
           Me.LN1.Font.Strikethrough = True
           Me.LN1.ForeColor = &H808080
           Else
           Me.LN1.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 1, "B").value
           Me.LN1.Font.Strikethrough = 0
           Me.LN1.ForeColor = &HFFFFFF
           End If
        Else
        Me.LN1.Caption = ""
        End If
        If AutoStart + 2 <= ARow Then
           If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 2, "F").value = 1 Then
           Me.LN2.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 2, "B").value & " - Done"
           Me.LN2.Font.Strikethrough = True
           Me.LN2.ForeColor = &H808080
           Else
           Me.LN2.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 2, "B").value
           Me.LN2.Font.Strikethrough = 0
           Me.LN2.ForeColor = &HFFFFFF
           End If
        Else
        Me.LN2.Caption = ""
        End If
        If AutoStart - 1 >= 1 Then
           If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 1, "F").value = 1 Then
           Me.LL1.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 1, "B").value
           Me.LL1.Font.Strikethrough = True
           Me.LL1.ForeColor = &H808080
           Else
           Me.LL1.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 1, "B").value
           Me.LL1.Font.Strikethrough = 0
           Me.LL1.ForeColor = &HFFFFFF
           End If
        Else
        Me.LL1.Caption = ""
        End If
        If AutoStart - 2 >= 1 Then
           If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 2, "F").value = 1 Then
           Me.LL2.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 2, "B").value
           Me.LL2.Font.Strikethrough = True
           Me.LL2.ForeColor = &H808080
           Else
           Me.LL2.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 2, "B").value
           Me.LL2.Font.Strikethrough = 0
           Me.LL2.ForeColor = &HFFFFFF
           End If
        Else
        Me.LL2.Caption = ""
        End If

If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "C").value <> "" And AUTO1 = 1 Then
   DisableEvents = True
   Me.Montant.value = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "C").value
   DisableEvents = False
   Me.Label6.ForeColor = &HFFFFC0
   Me.FrameAuto.ForeColor = &HFFFFC0
   If Me.Cat_1.value = SIncome Then
   Me.Cat_1.value = Expense
   End If
End If
If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "D").value <> "" Then
   DisableEvents = True
   Me.Montant.value = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "D").value
   DisableEvents = False
   Me.Label6.ForeColor = &HFFFFC0
   Me.FrameAuto.ForeColor = &HFFFFC0
   If Me.Cat_1.value = Expense Then
   Me.Cat_1.value = SIncome
   End If
End If
AutoM = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "E").value
FirstMonth = ThisWorkbook.Sheets("Library").Cells(1 + AutoM, "R").value
On Error Resume Next
Me.Mois.value = FirstMonth
On Error GoTo 0
End If
End Sub

Private Sub CommandDelete_Click()
Set helperRange = ThisWorkbook.Sheets("Input").Range("Z1:Z" & ThisWorkbook.Sheets("Input").Cells(ThisWorkbook.Sheets("Input").Rows.Count, "Z").End(xlUp).Row)

    ' Get the corresponding row in the helper range
    If Me.PPA.ListIndex = -1 Then
    Exit Sub
    End If
          If Me.PPA.ListIndex + 1 = Me.PPA.ListCount Then
          Rows = tbh.ListColumns("Max1").DataBodyRange(1, 1).value
               If tb.ListColumns("Année").DataBodyRange(1, 1).value = "Exemple" Then
               Exit Sub
               End If
   
          Set rng = tb.ListRows(Rows).Range
          rng.Delete
          Me.PPA.value = ""
             If Rows = 1 Then
             Me.Last_Cash.Caption = "---"
             tb.ListColumns("Année").Range(2, 1).value = "Exemple"
             Else

             Rows = Rows - 1
             Me.Last_Cash.Caption = tb.ListColumns("Montant").DataBodyRange(Rows, 1).value
             End If
               If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets("Input").Range("Z:Z")) = 0 Then
               selectedRow = 1 ' If column Z is empty, start at row 1
               Else
               selectedRow = ThisWorkbook.Sheets("Input").Cells(ThisWorkbook.Sheets("Input").Rows.Count, "Z").End(xlUp).Row
               End If
            helperRange.Cells(selectedRow, 1).value = ""
            Set helperRange = ThisWorkbook.Sheets("Input").Range("Z1:Z" & ThisWorkbook.Sheets("Input").Cells(ThisWorkbook.Sheets("Input").Rows.Count, "Z").End(xlUp).Row)
            Me.PPA.RowSource = "'" & helperRange.Worksheet.Name & "'!" & helperRange.Address
          
    Else
    SelectedValue = Me.PPA.value
    selectedRow = helperRange.Cells(Me.PPA.ListIndex + 1).Row
    helperRange.Cells(selectedRow, 1).value = "Deleted"
    Me.PPA.RowSource = "'" & helperRange.Worksheet.Name & "'!" & helperRange.Address
    selectedRow = Start + selectedRow - 2
    tb.ListColumns("Check").DataBodyRange(selectedRow, 1).value = 3
    End If
    Me.PPA.ListIndex = -1
    Me.PPA.TopIndex = Me.PPA.ListCount - 1
End Sub

Private Sub CommandHDelete_Click()
Set helperRange = ThisWorkbook.Sheets("Input").Range("Y1:Y" & ThisWorkbook.Sheets("Input").Cells(ThisWorkbook.Sheets("Input").Rows.Count, "Y").End(xlUp).Row)

    ' Get the corresponding row in the helper range
    If Me.Correction.ListIndex >= 0 Then
    SelectedValue = Me.Correction.value
    selectedRow = helperRange.Cells(Me.Correction.ListIndex + 1).Row
    helperRange.Cells(selectedRow, 1).value = helperRange.Cells(selectedRow, 1).value & " --- Deleted"
    Me.Correction.RowSource = "'" & helperRange.Worksheet.Name & "'!" & helperRange.Address
    Else
    Exit Sub
    End If
    selectedRow = selectedRow
    tb.ListColumns("Check").DataBodyRange(selectedRow, 1).value = 4
    V1 = tb.ListColumns("Année").DataBodyRange(selectedRow, 1).value
    V2 = tb.ListColumns("Mois").DataBodyRange(selectedRow, 1).value
    V3 = tb.ListColumns("Cat_1").DataBodyRange(selectedRow, 1).value
    V4 = tb.ListColumns("Cat_2").DataBodyRange(selectedRow, 1).value
    V5 = tb.ListColumns("Cat_3").DataBodyRange(selectedRow, 1).value
    V6 = tb.ListColumns("Montant").DataBodyRange(selectedRow, 1).value
    V6 = V6 * (-1)
    Last_Row = tb.ListRows.Count + 1
    tb.ListColumns("Année").DataBodyRange(Last_Row, 1).value = V1
    tb.ListColumns("Mois").DataBodyRange(Last_Row, 1).value = V2
    tb.ListColumns("Cat_1").DataBodyRange(Last_Row, 1).value = V3
    tb.ListColumns("Cat_2").DataBodyRange(Last_Row, 1).value = V4
    tb.ListColumns("Cat_3").DataBodyRange(Last_Row, 1).value = V5
    tb.ListColumns("Montant").DataBodyRange(Last_Row, 1).value = V6
    
    Set helperRange = ThisWorkbook.Sheets("Input").Range("Z1:Z" & ThisWorkbook.Sheets("Input").Cells(ThisWorkbook.Sheets("Input").Rows.Count, "Z").End(xlUp).Row)
    If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets("Input").Range("Z:Z")) = 0 Then
    selectedRow = 1 ' If column Z is empty, start at row 1
    Else
    selectedRow = ThisWorkbook.Sheets("Input").Cells(ThisWorkbook.Sheets("Input").Rows.Count, "Z").End(xlUp).Row
    selectedRow = selectedRow + 1
    End If
   SelectedValue = V1 & "-" & Left(V2, 3) & "-" & V5 & "-(" & V6 & CRY & ")"
   helperRange.Cells(selectedRow, 1).value = SelectedValue
   Set helperRange = ThisWorkbook.Sheets("Input").Range("Z1:Z" & ThisWorkbook.Sheets("Input").Cells(ThisWorkbook.Sheets("Input").Rows.Count, "Z").End(xlUp).Row)
   Me.PPA.RowSource = "'" & helperRange.Worksheet.Name & "'!" & helperRange.Address
   Me.PPA.TopIndex = Me.PPA.ListCount - 1
   Me.Correction.TopIndex = Me.Correction.ListCount - 1
    
End Sub


Private Sub CommandNext_Click()
AUTO1 = 1
If Tag = 1 Then
    
        AutoStart = AutoStart + 1 ' Default to the first row if no blank cells
        If AutoStart > ARow Then
        LibR = ThisWorkbook.Sheets("Library").Range("A219").value
        Me.FrameAuto.Caption = LibR
        Me.Label6.ForeColor = &HC0C0C0
        Me.FrameAuto.ForeColor = &HC0C0C0
        AUTO1 = 0
        AutoStart = AutoStart - 1
        Me.Montant.value = ""
        Me.CommandNext.Enabled = False
        Exit Sub
        End If
        Me.CommandButton1.Enabled = True
        If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "F").value = 1 Then
        Me.LAuto.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "B").value
        Me.LAuto.Font.Strikethrough = True
        Me.LAuto.ForeColor = &H808080
        LibR = ThisWorkbook.Sheets("Library").Range("A219").value
        Me.FrameAuto.Caption = LibR
        Me.FrameAuto.ForeColor = &HC0C0C0
        Me.Label6.ForeColor = &HC0C0C0
        Me.Montant.value = ""
        AUTO1 = 0
        Else
        Me.LAuto.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "B").value
        Me.LAuto.Font.Strikethrough = 0
        Me.LAuto.ForeColor = &HFFFFC0
        LibR = ThisWorkbook.Sheets("Library").Range("A212").value
        Me.FrameAuto.Caption = LibR
        AUTO1 = 1
        End If
        If AutoStart + 1 <= ARow Then
           If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 1, "F").value = 1 Then
           Me.LN1.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 1, "B").value
           Me.LN1.Font.Strikethrough = True
           Me.LN1.ForeColor = &H808080
           Else
           Me.LN1.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 1, "B").value
           Me.LN1.Font.Strikethrough = 0
           Me.LN1.ForeColor = &HFFFFFF
           End If
        Else
        Me.LN1.Caption = ""
        End If
        If AutoStart + 2 <= ARow Then
           If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 2, "F").value = 1 Then
           Me.LN2.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 2, "B").value
           Me.LN2.Font.Strikethrough = True
           Me.LN2.ForeColor = &H808080
           Else
           Me.LN2.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 2, "B").value
           Me.LN2.Font.Strikethrough = 0
           Me.LN2.ForeColor = &HFFFFFF
           End If
        Else
        Me.LN2.Caption = ""
        End If
        If AutoStart - 1 >= 1 Then
           If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 1, "F").value = 1 Then
           Me.LL1.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 1, "B").value
           Me.LL1.Font.Strikethrough = True
           Me.LL1.ForeColor = &H808080
           Else
           Me.LL1.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 1, "B").value
           Me.LL1.Font.Strikethrough = 0
           Me.LL1.ForeColor = &HFFFFFF
           End If
        Else
        Me.LL1.Caption = ""
        End If
        If AutoStart - 2 >= 1 Then
           If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 2, "F").value = 1 Then
           Me.LL2.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 2, "B").value
           Me.LL2.Font.Strikethrough = True
           Me.LL2.ForeColor = &H808080
           Else
           Me.LL2.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 2, "B").value
           Me.LL2.Font.Strikethrough = 0
           Me.LL2.ForeColor = &HFFFFFF
           End If
        Else
        Me.LL2.Caption = ""
        End If
Me.Cat_2.ListIndex = -1
Me.Cat_3.ListIndex = -1
If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "C").value <> "" And AUTO1 = 1 Then
   DisableEvents = True
   Me.Montant.value = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "C").value
   DisableEvents = False
   Me.Label6.ForeColor = &HFFFFC0
   Me.FrameAuto.ForeColor = &HFFFFC0
   Me.Cat_1.value = Expense
End If
If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "D").value <> "" Then
   DisableEvents = True
   Me.Montant.value = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "D").value
   DisableEvents = False
   Me.Label6.ForeColor = &HFFFFC0
   Me.FrameAuto.ForeColor = &HFFFFC0
   Me.Cat_1.value = SIncome
End If
AutoM = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "E").value
FirstMonth = ThisWorkbook.Sheets("Library").Cells(1 + AutoM, "R").value
On Error Resume Next
Me.Mois.value = FirstMonth
On Error GoTo 0
End If
End Sub

Private Sub Fermer_Click()
Unload Me
End Sub

Private Sub Mois_Click()
If Me.Mois.ListIndex >= 0 Then
Me.Mois.BorderStyle = fmBorderStyleNone
Me.PPA.value = ""
End If
End Sub

Private Sub Montant_Change()
Me.Montant.BorderStyle = fmBorderStyleNone
Me.Montant.ForeColor = vbWhite
Me.PPA.value = ""
If DisableEvents Then Exit Sub
Me.Label6.ForeColor = &HC0C0C0
Me.FrameAuto.ForeColor = &HC0C0C0
End Sub

Private Sub MultiPage1_Change()

Dim emptyCells As Range
Dim dataArray1(), dataArray2(), dataArray3(), dataArray4(), dataArray5(), dataArray6(), dataArray7(), dataDELETE() As Variant
If A = 1 Then
Exit Sub
End If
If Tag = 1 Then
If Me.MultiPage1.value = 2 Then
Me.FrameAuto.Visible = False
Me.CommandButton1.Enabled = False
Me.CommandNext.Enabled = False
ElseIf Me.MultiPage1.value = 1 Then
Me.FrameAuto.Visible = True
Me.CommandButton1.Enabled = True
Me.CommandNext.Enabled = True
End If
End If

On Error Resume Next
    Set Sarray = tb.ListColumns("Check").DataBodyRange
    Set emptyCells = Sarray.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0

    ' Determine the starting row
    If Not emptyCells Is Nothing And tb.ListColumns("Année").DataBodyRange.Cells(1, 1).value <> "Exemple" Then
        Set FindFirstEmptyCellInColumn = emptyCells.Cells(1)
        EndE = FindFirstEmptyCellInColumn.Row - 2
    Else
        Exit Sub ' Default to the first row if no blank cells
    End If

    ' Define end row and calculate range
    Start = 1
    Me.MousePointer = fmMousePointerHourGlass
    LibR = ThisWorkbook.Sheets("Library").Range("A220").value
    Me.Caption = LibR
    ' Initialize arrays
    ReDim dataArray1(1 To EndE - Start + 1)
    ReDim dataArray2(1 To EndE - Start + 1)
    ReDim dataArray3(1 To EndE - Start + 1)
    ReDim dataArray4(1 To EndE - Start + 1)
    ReDim dataArray5(1 To EndE - Start + 1)
    ReDim dataArray6(1 To EndE - Start + 1)
    ReDim dataArray7(1 To EndE - Start + 1)
    ReDim dataDELETE(1 To EndE - Start + 1)

    ' Populate arrays
    arrIndex = 1
    For i = Start To EndE
        dataArray1(arrIndex) = tb.DataBodyRange.Cells(i, tb.ListColumns("Année").Index).value
        dataArray2(arrIndex) = tb.DataBodyRange.Cells(i, tb.ListColumns("Mois").Index).value
        dataArray3(arrIndex) = tb.DataBodyRange.Cells(i, tb.ListColumns("Cat_1").Index).value
        dataArray4(arrIndex) = tb.DataBodyRange.Cells(i, tb.ListColumns("Cat_2").Index).value
        dataArray5(arrIndex) = tb.DataBodyRange.Cells(i, tb.ListColumns("Cat_3").Index).value
        dataArray6(arrIndex) = tb.DataBodyRange.Cells(i, tb.ListColumns("Montant").Index).value
        If tb.ListColumns("Check").DataBodyRange(i, 1).value = 4 Then
        dataArray7(arrIndex) = " --- Deleted"
        Else
        dataArray7(arrIndex) = " "
        End If
        ' Concatenate data into dataMaster
        dataDELETE(arrIndex) = dataArray1(arrIndex) & " - " & dataArray2(arrIndex) & " - " & dataArray3(arrIndex) & " - " & dataArray4(arrIndex) & " - " & dataArray5(arrIndex) & " - ( " & dataArray6(arrIndex) & CRY & " )" & dataArray7(arrIndex)

        arrIndex = arrIndex + 1
    Next i
    
    ' Clear any existing data
    ThisWorkbook.Sheets("Input").Range("Y1:Y" & ThisWorkbook.Sheets("Input").Rows.Count).ClearContents
    
    ' Write the dataMaster array to the helper range
    For i = LBound(dataDELETE) To UBound(dataDELETE)
        ThisWorkbook.Sheets("Input").Cells(i, 25).value = dataDELETE(i)
    Next i
    
    ' Set the range for RowSource
    Set helperRange = ThisWorkbook.Sheets("Input").Range("Y1:Y" & UBound(dataDELETE))


' Set the RowSource of the ComboBox (PPA)
Me.Correction.RowSource = "'" & helperRange.Worksheet.Name & "'!" & helperRange.Address
Me.Correction.TopIndex = Me.Correction.ListCount - 1
A = 1
   Me.MousePointer = fmMousePointerDefault
   LibR = ThisWorkbook.Sheets("Library").Range("A96").value
   Me.Caption = LibR
End Sub

Private Sub Test_Click()
Dim LibR, LibR1, LibR2 As String

ER = 0
If Me.Montant.value = "" Then
   Me.Montant.BorderStyle = fmBorderStyleSingle
   Me.Montant.BorderColor = vbRed
   ER = 1
Else
    VMontant = Replace(Me.Montant.value, localDec, altDec)
    If vbaSep = "." Then
    VMontant = Replace(VMontant, altDec, Application.International(xlDecimalSeparator))
    End If
    VMontant = Trim(VMontant)
    If IsNumeric(VMontant) Then
        Cash = CDbl(VMontant)
    Else
    Me.Montant.ForeColor = vbRed
    ER = 1
    End If
End If

Rows = tbh.ListColumns("Max1").DataBodyRange(1, 1).value + 1

    
 If Me.Year.value = "" Then
    Me.Year.BorderStyle = fmBorderStyleSingle
    Me.Year.BorderColor = vbRed
    ER = 1
End If

If Me.Cat_2.ListIndex = -1 Then
    Me.LCAT2.BorderStyle = fmBorderStyleSingle
    Me.LCAT2.BorderColor = vbRed
    ER = 1
End If

If Me.Cat_3.ListIndex = -1 Then
    Me.LCAT3.BorderStyle = fmBorderStyleSingle
    Me.LCAT3.BorderColor = vbRed
    ER = 1
End If

If ER = 1 Then
   Exit Sub
End If
    
If Me.Year.value <> "" And Me.Mois.ListIndex >= 0 And Me.Cat_1.ListIndex >= 0 And Me.Cat_2.ListIndex >= 0 And Me.Cat_3.ListIndex >= 0 Then
   tb.ListColumns("Année").DataBodyRange(Rows, 1).value = Me.Year.value
   tb.ListColumns("Mois").DataBodyRange(Rows, 1).value = Me.Mois.value
   tb.ListColumns("Cat_1").DataBodyRange(Rows, 1).value = Me.Cat_1.value
   tb.ListColumns("Cat_2").DataBodyRange(Rows, 1).value = Me.Cat_2.value
   tb.ListColumns("Cat_3").DataBodyRange(Rows, 1).value = Me.Cat_3.value
   LibR1 = ThisWorkbook.Sheets("Library").Range("A246").value
   LibR2 = ThisWorkbook.Sheets("Library").Range("A247").value
   Select Case Me.Test.Caption
          Case Is = LibR1
               If Cash < 0 Then
                  Cash = Cash * -1
               End If
          Case Is = LibR2
              If Cash > 0 Then
                  Cash = Cash * -1
               End If
   End Select
   tb.ListColumns("Montant").DataBodyRange(Rows, 1).value = Cash
   Me.Last_Cash.Caption = Cash
   Me.Montant.value = ""
   Me.PPA.value = ""
   
   If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets("Input").Range("Z:Z")) = 0 Then
    selectedRow = 1 ' If column Z is empty, start at row 1
    Else
    selectedRow = ThisWorkbook.Sheets("Input").Cells(ThisWorkbook.Sheets("Input").Rows.Count, "Z").End(xlUp).Row
    selectedRow = selectedRow + 1
    End If
   SelectedValue = Me.Year.value & "-" & Left(Me.Mois.value, 3) & "-" & Me.Cat_3.value & "-(" & Cash & CRY & ")"
   helperRange.Cells(selectedRow, 1).value = SelectedValue
   Set helperRange = ThisWorkbook.Sheets("Input").Range("Z1:Z" & ThisWorkbook.Sheets("Input").Cells(ThisWorkbook.Sheets("Input").Rows.Count, "Z").End(xlUp).Row)
   Me.PPA.RowSource = "'" & helperRange.Worksheet.Name & "'!" & helperRange.Address
   Me.PPA.TopIndex = Me.PPA.ListCount - 1
Else
LibR = ThisWorkbook.Sheets("Library").Range("A104").value
MsgBox (LibR), vbOKOnly
End If

If Tag = 1 And AUTO1 = 1 Then
        
        ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "F").value = 1
        AutoStart = AutoStart + 1 ' Default to the first row if no blank cells
        If AutoStart > ARow Then
        Me.LAuto.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 1, "B").value
        Me.LAuto.Font.Strikethrough = True
        Me.LAuto.ForeColor = &H808080
        LibR = ThisWorkbook.Sheets("Library").Range("A219").value
        Me.FrameAuto.Caption = LibR
        Me.FrameAuto.ForeColor = &HC0C0C0
        AUTO1 = 0
        Me.Montant.value = ""
        Me.Label6.ForeColor = &HC0C0C0
        Me.CommandNext.Enabled = False
        Exit Sub
        End If
        Me.CommandButton1.Enabled = True
        If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "F").value = 1 Then
        Me.LAuto.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "B").value
        Me.LAuto.Font.Strikethrough = True
        Me.LAuto.ForeColor = &H808080
        LibR = ThisWorkbook.Sheets("Library").Range("A219").value
        Me.FrameAuto.Caption = LibR
        Me.FrameAuto.ForeColor = &HC0C0C0
        Me.Label6.ForeColor = &HC0C0C0
        Me.Montant.value = ""
        AUTO1 = 0
        Else
        Me.LAuto.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "B").value
        Me.LAuto.Font.Strikethrough = 0
        Me.LAuto.ForeColor = &HFFFFC0
        LibR = ThisWorkbook.Sheets("Library").Range("A212").value
        Me.FrameAuto.Caption = LibR
        AUTO1 = 1
        End If
        If AutoStart + 1 <= ARow Then
           If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 1, "F").value = 1 Then
           Me.LN1.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 1, "B").value
           Me.LN1.Font.Strikethrough = True
           Me.LN1.ForeColor = &H808080
           Else
           Me.LN1.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 1, "B").value
           Me.LN1.Font.Strikethrough = 0
           Me.LN1.ForeColor = &HFFFFFF
           End If
        Else
        Me.LN1.Caption = ""
        End If
        If AutoStart + 2 <= ARow Then
           If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 2, "F").value = 1 Then
           Me.LN2.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 2, "B").value
           Me.LN2.Font.Strikethrough = True
           Me.LN2.ForeColor = &H808080
           Else
           Me.LN2.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 2, "B").value
           Me.LN2.Font.Strikethrough = 0
           Me.LN2.ForeColor = &HFFFFFF
           End If
        Else
        Me.LN2.Caption = ""
        End If
        If AutoStart - 1 >= 1 Then
           If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 1, "F").value = 1 Then
           Me.LL1.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 1, "B").value
           Me.LL1.Font.Strikethrough = True
           Me.LL1.ForeColor = &H808080
           Else
           Me.LL1.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 1, "B").value
           Me.LL1.Font.Strikethrough = 0
           Me.LL1.ForeColor = &HFFFFFF
           End If
        Else
        Me.LL1.Caption = ""
        End If
        If AutoStart - 2 >= 1 Then
           If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 2, "F").value = 1 Then
           Me.LL2.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 2, "B").value
           Me.LL2.Font.Strikethrough = True
           Me.LL2.ForeColor = &H808080
           Else
           Me.LL2.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 2, "B").value
           Me.LL2.Font.Strikethrough = 0
           Me.LL2.ForeColor = &HFFFFFF
           End If
        Else
        Me.LL2.Caption = ""
        End If

If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "C").value <> "" And AUTO1 = 1 Then
   DisableEvents = True
   Me.Montant.value = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "C").value
   DisableEvents = False
   Me.Label6.ForeColor = &HFFFFC0
   Me.FrameAuto.ForeColor = &HFFFFC0
   If Me.Cat_1.value = SIncome Then
   Me.Cat_1.value = Expense
   End If
End If
If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "D").value <> "" Then
   DisableEvents = True
   Me.Montant.value = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "D").value
   DisableEvents = False
   Me.Label6.ForeColor = &HFFFFC0
   Me.FrameAuto.ForeColor = &HFFFFC0
   If Me.Cat_1.value = Expense Then
   Me.Cat_1.value = SIncome
   End If
End If
AutoM = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "E").value
FirstMonth = ThisWorkbook.Sheets("Library").Cells(1 + AutoM, "R").value
On Error Resume Next
Me.Mois.value = FirstMonth
On Error GoTo 0
End If
End Sub
Private Sub UserForm_Initialize()
Dim An As Integer
Dim FirstMonth, LibR As String
Dim emptyCells As Range
Dim dataArray1(), dataArray2(), dataArray3(), dataArray4(), dataArray5(), dataArray6(), dataMaster() As Variant

SetUserFormBorderColor Me, RGB(0, 0, 0)

 localDec = Application.International(xlDecimalSeparator)
    
    ' Define the alternative decimal separator: if local is dot then alt is comma; otherwise, alt is dot.
    If localDec = "." Then
        altDec = ","
    Else
        altDec = "."
    End If
    
vbaSep = Mid$(Format(1.1, "0.0"), 2, 1)

Me.Cat_1.ListIndex = -1
AUTO1 = 0
NewOne = 0
Tag = 0
LibR = ThisWorkbook.Sheets("Library").Range("A96").value
Me.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A97").value
Me.LCAT2.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A54").value
Me.LCAT3.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A99").value
Me.Fermer.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A100").value
Me.CommandDelete.Caption = LibR
Me.CommandHDelete.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A101").value
Me.Label6.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A102").value
Me.Label8.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A103").value
Me.Cancel.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A105").value
Me.MultiPage1.Pages(0).Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A106").value
Me.MultiPage1.Pages(1).Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A212").value
Me.FrameAuto.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A242").value
Me.CommandNext.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A241").value
Me.CommandButton1.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A243").value
Me.Cancel.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A244").value
Me.CommandDelete.Caption = LibR
Me.CommandHDelete.Caption = LibR
CodeM = ThisWorkbook.Sheets("Library").Range("R14").value
LibR = ThisWorkbook.Sheets("Library").Range("A245").value
Me.Add.Caption = LibR

Set tbRefin = ThisWorkbook.Sheets("WB_Data").ListObjects("TbREFIncome")
Set tbCCY = ThisWorkbook.Sheets("WB_Data").ListObjects("tbCCY")
If tbCCY.ListColumns("Path").DataBodyRange(4, 1).value = 1 Then
   Tag = 1
End If
CRY = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
Me.SIG1.Caption = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
Me.SIG2.Caption = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
Set tbh = ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead")
Rows = tbh.ListColumns("Max1").DataBodyRange(1, 1).value
Set tb = ThisWorkbook.Sheets("Input").ListObjects("DataTable")
An = VBA.DateTime.Year(Date)
Me.Year.value = An
FirstMonth = Application.WorksheetFunction.Text(Date, CodeM & "mmmm")
Me.Mois.value = FirstMonth
Expense = tbRefin.ListColumns("SuperCat").DataBodyRange(1, 1).value
SIncome = tbRefin.ListColumns("SuperCat").DataBodyRange(2, 1).value
Me.Cat_1.value = Expense
Me.Additional.Caption = Chr(43)
Savings = ThisWorkbook.Sheets("WB_Data").ListObjects("Income").ListColumns("C22").DataBodyRange(2, 1).value
A = 0

   If tb.ListColumns("Année").DataBodyRange(1, 1).value <> "Exemple" Then
      Me.Last_Cash.Caption = tb.ListColumns("Montant").DataBodyRange(Rows, 1).value
      Rows = Rows + 1
   ElseIf Rows = 0 Then
      Rows = 1
      Me.Last_Cash.Caption = "--"
   End If
   

On Error Resume Next
    Set Sarray = tb.ListColumns("Check").DataBodyRange
    Set emptyCells = Sarray.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0

    ' Determine the starting row
    If Not emptyCells Is Nothing And tb.ListColumns("Année").DataBodyRange.Cells(1, 1).value <> "Exemple" Then
        Set FindFirstEmptyCellInColumn = emptyCells.Cells(1)
        Start = FindFirstEmptyCellInColumn.Row
    Else
        Start = 1 ' Default to the first row if no blank cells
        NewOne = 1
        ThisWorkbook.Sheets("Input").Range("Z1:Z" & ThisWorkbook.Sheets("Input").Rows.Count).ClearContents
        Set helperRange = ThisWorkbook.Sheets("Input").Range("Z1:Z" & ThisWorkbook.Sheets("Input").Cells(ThisWorkbook.Sheets("Input").Rows.Count, "Z").End(xlUp).Row)
    End If
    
    If NewOne = 0 Then
    ' Define end row and calculate range
    EndE = tb.ListRows.Count + 1

    ' Initialize arrays
    ReDim dataArray1(1 To EndE - Start + 1)
    ReDim dataArray2(1 To EndE - Start + 1)
    ReDim dataArray3(1 To EndE - Start + 1)
    ReDim dataArray4(1 To EndE - Start + 1)
    ReDim dataArray5(1 To EndE - Start + 1)
    ReDim dataArray6(1 To EndE - Start + 1)
    ReDim dataMaster(1 To EndE - Start + 1)

    ' Populate arrays
    arrIndex = 1
    For i = Start To EndE
        If tb.ListColumns("Check").DataBodyRange(i - 1, 1).value = 3 Then
        dataMaster(arrIndex) = "Deleted"
        Else
        dataArray1(arrIndex) = tb.DataBodyRange.Cells(i - tb.DataBodyRange.Row + 1, tb.ListColumns("Année").Index).value
        dataArray2(arrIndex) = Left(tb.DataBodyRange.Cells(i - tb.DataBodyRange.Row + 1, tb.ListColumns("Mois").Index).value, 3)
        dataArray5(arrIndex) = tb.DataBodyRange.Cells(i - tb.DataBodyRange.Row + 1, tb.ListColumns("Cat_3").Index).value
        dataArray6(arrIndex) = tb.DataBodyRange.Cells(i - tb.DataBodyRange.Row + 1, tb.ListColumns("Montant").Index).value

        ' Concatenate data into dataMaster
        dataMaster(arrIndex) = dataArray1(arrIndex) & "-" & dataArray2(arrIndex) & "-" & _
                               dataArray5(arrIndex) & "-(" & dataArray6(arrIndex) & CRY & ")"
        End If
        arrIndex = arrIndex + 1
    Next i
  
    ' Clear any existing data
    ThisWorkbook.Sheets("Input").Range("Z1:Z" & ThisWorkbook.Sheets("Input").Rows.Count).ClearContents

    ' Write the dataMaster array to the helper range
    For i = LBound(dataMaster) To UBound(dataMaster)
        ThisWorkbook.Sheets("Input").Cells(i, 26).value = dataMaster(i)
    Next i

    ' Set the range for RowSource
    Set helperRange = ThisWorkbook.Sheets("Input").Range("Z1:Z" & UBound(dataMaster))

' Set the RowSource of the ComboBox (PPA)

Me.PPA.RowSource = "'" & helperRange.Worksheet.Name & "'!" & helperRange.Address
Me.PPA.TopIndex = Me.PPA.ListCount - 1

   End If
If Tag = 1 Then
   Me.FrameAuto.Visible = True
   Me.CommandButton1.Enabled = True
   Me.CommandNext.Enabled = True
   ARow = ThisWorkbook.Sheets("Data Entry").Cells(ThisWorkbook.Sheets("Data Entry").Rows.Count, "B").End(xlUp).Row
   On Error Resume Next
    Set Autoarray = ThisWorkbook.Sheets("Data Entry").Range(ThisWorkbook.Sheets("Data Entry").Cells(1, "F"), ThisWorkbook.Sheets("Data Entry").Cells(ARow, "F"))
    Set emptyCells = Autoarray.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0

    ' Determine the starting row
    If Not emptyCells Is Nothing Then
        Set FindFirstEmptyCellInColumn = emptyCells.Cells(1)
        AutoStart = FindFirstEmptyCellInColumn.Row
    Else
        AutoStart = 1 ' Default to the first row if no blank cells
    End If
        If AutoStart = 1 Then
           Me.CommandButton1.Enabled = False
        ElseIf AutoStart = ARow Then
           Me.CommandNext.Enabled = False
        End If
        If ARow = 1 Then
              Me.CommandNext.Enabled = False
              Me.CommandButton1.Enabled = False
        End If
        If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "F").value = 1 Then
        AUTO1 = 0
        LibR = ThisWorkbook.Sheets("Library").Range("A219").value
        Me.FrameAuto.Caption = LibR
        Me.LAuto.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "B").value
        Me.LAuto.Font.Strikethrough = True
        Me.LAuto.ForeColor = &H808080
        Else
        AUTO1 = 1
        LibR = ThisWorkbook.Sheets("Library").Range("A212").value
        Me.FrameAuto.Caption = LibR
        Me.LAuto.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "B").value
        Me.LAuto.Font.Strikethrough = 0
        End If
        If AutoStart + 1 <= ARow Then
           If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 1, "F").value = 1 Then
           Me.LN1.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 1, "B").value
           Me.LN1.Font.Strikethrough = True
           Me.LN1.ForeColor = &H808080
           Else
           Me.LN1.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 1, "B").value
           Me.LN1.ForeColor = &HFFFFFF
           Me.LN1.Font.Strikethrough = 0
           End If
        Else
        Me.LN1.Caption = ""
        End If
        If AutoStart + 2 <= ARow Then
           If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 2, "F").value = 1 Then
           Me.LN2.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 2, "B").value
           Me.LN2.Font.Strikethrough = True
           Me.LN2.ForeColor = &H808080
           Else
           Me.LN2.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart + 2, "B").value
           Me.LN2.Font.Strikethrough = 0
           Me.LN2.ForeColor = &HFFFFFF
           End If
        Else
        Me.LN2.Caption = ""
        End If
        If AutoStart - 1 >= 1 Then
           If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 1, "F").value = 1 Then
           Me.LL1.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 1, "B").value
           Me.LL1.Font.Strikethrough = True
           Me.LL1.ForeColor = &H808080
           Else
           Me.LL1.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 1, "B").value
           Me.LL1.Font.Strikethrough = 0
           Me.LL1.ForeColor = &HFFFFFF
           End If
        Else
        Me.LL1.Caption = ""
        End If
        If AutoStart - 2 >= 1 Then
           If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 2, "F").value = 1 Then
           Me.LL2.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 2, "B").value
           Me.LL2.Font.Strikethrough = True
           Me.LL2.ForeColor = &H808080
           Else
           Me.LL2.Caption = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart - 2, "B").value
           Me.LL1.Font.Strikethrough = 0
           Me.LL2.ForeColor = &HFFFFFF
           End If
        Else
        Me.LL2.Caption = ""
        End If
     
If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "C").value <> "" And AUTO1 = 1 Then
   DisableEvents = True
   Me.Montant.value = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "C").value
   DisableEvents = False
   Me.Label6.ForeColor = &HFFFFC0
   Me.FrameAuto.ForeColor = &HFFFFC0
   If Me.Cat_1.value = SIncome Then
   Me.Cat_1.value = Expense
   End If
End If
If ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "D").value <> "" Then
   DisableEvents = True
   Me.Montant.value = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "D").value
   DisableEvents = False
   Me.Label6.ForeColor = &HFFFFC0
   Me.FrameAuto.ForeColor = &HFFFFC0
   If Me.Cat_1.value = Expense Then
   Me.Cat_1.value = SIncome
   End If
End If
AutoM = ThisWorkbook.Sheets("Data Entry").Cells(AutoStart, "E").value
On Error Resume Next
FirstMonth = ThisWorkbook.Sheets("Library").Cells(1 + AutoM, "R").value
Me.Mois.value = FirstMonth
On Error GoTo 0


End If


End Sub

Private Sub Year_Change()

End Sub
