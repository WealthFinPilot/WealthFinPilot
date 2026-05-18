Attribute VB_Name = "Automate"
Public tbCCY As ListObject
Public ER As Integer
Public localDec, altDec, vbaSep As String
Sub ExtractTransactionDataWithDeletion()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Data Entry")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim transactionMonth As Integer
    Dim i As Long, outputRow As Long
    Dim data As Variant
    Dim lineNumber As Long, transactionDate As String, description As String, CRY As String
    Dim value7 As Variant, value8 As Variant, extractedValue As Variant
    Dim skippedCount As Long: skippedCount = 0
    Dim CodeM As String
    Set tbCCY = ThisWorkbook.Sheets("WB_Data").ListObjects("tbCCY")
    CodeM = ThisWorkbook.Sheets("Library").Range("R14").value
    CRY = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
    outputRow = 1 ' Initialize output for Column E

    If ws.Range("C1").value <> "" And ws.Range("D1").value <> "" And ws.Range("E1").value <> "" And ws.Range("F1").value <> "" Then
       lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
           For i = lastRow To 1 Step -1
               lineNumber = CLng(ws.Cells(i, "E").value)
               transactionDate = Format(CDate(Trim(ws.Cells(i, "D").value)), CodeM & "dd mmm") ' Data4: Date (formatted)
               transactionMonth = Month(CDate(Trim(ws.Cells(i, "D").value)))
               description = ws.Cells(i, "F").value ' Data6: Description
               value7 = SafeConvertToNumber(ws.Cells(i, "H").value, True)  ' Withdrawal (Negative if present)
               value8 = SafeConvertToNumber(ws.Cells(i, "I").value, False) ' Revenue (Positive if present)
               
                If Left(description, 23) = "Paiement facture - Bank" And Right(description, 5) = "Bank1" Then
                   ws.Rows(i).Delete
                   skippedCount = skippedCount + 1
                   GoTo NextLine
                End If
        
               If ws.Cells(i, "C").value <> "EOP" Then
                  ws.Rows(i).Delete
                  skippedCount = skippedCount + 1
                  GoTo NextLine
               Else
               ws.Cells(i, "C").value = ""
               End If
               
               
                extractedValue = IIf(value7 <> 0, value7, value8)
        
        '    Output formatted data in Column B
             ws.Cells(i, "B").value = transactionDate & " - " & description & " / (" & extractedValue & CRY & ")"
        
        ' Handle "Virement Interac" special case
        If Left(description, 19) = "Virement Interac de" Then
            ws.Cells(i, "C").value = extractedValue ' Revenue in Column C
            ws.Cells(i, "D").value = ""
        ElseIf value7 <> 0 Then
            ws.Cells(i, "C").value = value7 ' Withdrawal (negative) in Column C
            ws.Cells(i, "D").value = ""
        ElseIf value8 <> 0 Then
            ws.Cells(i, "D").value = value8 ' Revenue (positive) in Column D
            ws.Cells(i, "C").value = ""
        End If
        ws.Cells(i, "E").value = transactionMonth
NextLine:
    Next i
    
    For i = 1 To lastRow - skippedCount
    ws.Cells(i, "B").value = "Line " & Format(i, "00000") & " - " & ws.Cells(i, "B").value
    Next i
    ws.Columns("F:F").ClearContents
    ws.Columns("G:G").ClearContents
    ws.Columns("H:H").ClearContents
    ws.Columns("I:I").ClearContents
    ws.Columns("N:N").ClearContents
    ws.Columns("C:C").NumberFormat = "General"
    ws.Columns("D:D").NumberFormat = "General"
               
    
    Else
    ' Loop from bottom to top to avoid skipping rows during deletion
    For i = lastRow To 1 Step -1
        ' Parse data from CSV in column A
        data = SplitCSV(ws.Cells(i, "A").value)
        
        ' Ensure valid data length (skip if malformed)
        If UBound(data) < 8 Then GoTo NextRow
        
        ' Extract core data
        lineNumber = CLng(data(4)) ' Data5: Line number
        transactionDate = Format(CDate(Trim(data(3))), CodeM & "dd mmm") ' Data4: Date (formatted)
        transactionMonth = Month(CDate(Trim(data(3))))
        description = Trim(data(5)) ' Data6: Description
        
        ' ? Safely convert Data7 and Data8 (handle empty fields)
        value7 = SafeConvertToNumber(data(7), True)  ' Withdrawal (Negative if present)
        value8 = SafeConvertToNumber(data(8), False) ' Revenue (Positive if present)

        ' Delete row if description starts with "Paiement facture - Bank"
        If Left(description, 23) = "Paiement facture - Bank" And Right(description, 5) = "Bank1" Then
            ws.Rows(i).Delete
            skippedCount = skippedCount + 1
            GoTo NextRow
        End If
        
        If data(2) <> "EOP" Then
            ws.Rows(i).Delete
            skippedCount = skippedCount + 1
            GoTo NextRow
        End If

        
        ' Choose the primary amount to display (Data9 equivalent)
        extractedValue = IIf(value7 <> 0, value7, value8)
        
        ' Output formatted data in Column B
        ws.Cells(i, "B").value = transactionDate & " - " & description & " / (" & extractedValue & CRY & ")"
        
        ' Handle "Virement Interac" special case
        If Left(description, 19) = "Virement Interac de" Then
            ws.Cells(i, "C").value = extractedValue ' Revenue in Column C
        ElseIf value7 <> 0 Then
            ws.Cells(i, "C").value = value7 ' Withdrawal (negative) in Column C
        ElseIf value8 <> 0 Then
            ws.Cells(i, "D").value = value8 ' Revenue (positive) in Column D
        End If
        ws.Cells(i, "E").value = transactionMonth
NextRow:
    Next i
    
    For i = 1 To lastRow - skippedCount
    ws.Cells(i, "B").value = "Line " & Format(i, "00000") & " - " & ws.Cells(i, "B").value
    Next i
    End If
End Sub

' Helper function to parse CSV while handling quotes
Function SplitCSV(ByVal csv As String) As Variant
    Dim parts As Object: Set parts = CreateObject("Scripting.Dictionary")
    Dim buffer As String: buffer = ""
    Dim inQuotes As Boolean: inQuotes = False
    Dim i As Long
    
    ' Parse CSV character by character
    For i = 1 To Len(csv)
        Dim char As String: char = Mid(csv, i, 1)
        
        If char = """" Then
            inQuotes = Not inQuotes
        ElseIf char = "," And Not inQuotes Then
            parts(parts.Count) = Trim(buffer)
            buffer = ""
        Else
            buffer = buffer & char
        End If
    Next i
    
    ' Capture final part
    If Len(buffer) > 0 Then parts(parts.Count) = Trim(buffer)
    SplitCSV = parts.Items
End Function

' Helper function to safely convert a CSV value to a number
Function SafeConvertToNumber(ByVal value As String, ByVal isNegative As Boolean) As Double
    ' Remove surrounding quotes if present
    If Left(value, 1) = """" And Right(value, 1) = """" Then
        value = Mid(value, 2, Len(value) - 2)
    End If
    
    ' Replace dot with comma if necessary (for European locales)
    If InStr(1, value, altDec) > 0 And Application.DecimalSeparator = altDec Then
        value = Replace(value, altDec, localDec)
    End If
    
    ' Convert to number if it's numeric
    If IsNumeric(value) Then
        If isNegative Then
            SafeConvertToNumber = -CDbl(value)
        Else
            SafeConvertToNumber = CDbl(value)
        End If
    Else
        SafeConvertToNumber = 0
    End If
End Function
Sub Choise_Module()
Application.ScreenUpdating = False
AutoAsk.Show
Set tbCCY = ThisWorkbook.Sheets("WB_Data").ListObjects("tbCCY")
ER = 0

localDec = Application.International(xlDecimalSeparator)
    
    ' Define the alternative decimal separator: if local is dot then alt is comma; otherwise, alt is dot.
    If localDec = "." Then
        altDec = ","
    Else
        altDec = "."
    End If
vbaSep = Mid$(Format(1.1, "0.0"), 2, 1)



  ThisWorkbook.Sheets("Data Entry").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
  If tbCCY.ListColumns("Path").DataBodyRange(4, 1).value = "" Then
   ThisWorkbook.Sheets("Data Entry").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
   ThisWorkbook.Sheets("DashBoard").Activate
   Application.ScreenUpdating = True
   Exit Sub
  ElseIf tbCCY.ListColumns("Path").DataBodyRange(4, 1).value = 3 Then
   Reset_Auto
   tbCCY.ListColumns("Path").DataBodyRange(4, 1).value = ""
   ThisWorkbook.Sheets("Data Entry").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
   ThisWorkbook.Sheets("DashBoard").Activate
   Application.ScreenUpdating = True
   Exit Sub
  ElseIf tbCCY.ListColumns("Path").DataBodyRange(4, 1).value = 1 Then
        ThisWorkbook.Sheets("Data Entry").Columns("E:E").ClearContents
        ThisWorkbook.Sheets("Data Entry").Columns("F:F").ClearContents
        If tbCCY.ListColumns("Path").DataBodyRange(3, 1).value = 1 Then
           Select Case tbCCY.ListColumns("SelectBank").DataBodyRange(1, 1).value
           Case Is = "Bank1"
           CopySheetFromWorkbook
           Case Is = "Bank2"
           SG
           Case Is = "Bank3"
           Case Is = ""
           Exit Sub
           Case Else
           Exit Sub
           End Select
        ElseIf tbCCY.ListColumns("Path").DataBodyRange(3, 1).value = 2 Then
        CREDIT
        ElseIf tbCCY.ListColumns("Path").DataBodyRange(3, 1).value = 3 Then
        PARDEFAUT
        End If
  ElseIf tbCCY.ListColumns("Path").DataBodyRange(4, 1).value = 2 Then
  tbCCY.ListColumns("Path").DataBodyRange(4, 1).value = 1
  End If
           If ThisWorkbook.Sheets("Data Entry").Range("A1").value <> "" And ThisWorkbook.Sheets("Data Entry").Range("B1").value <> "" And ER = 0 Then
           Chart_Analyse.AUTOMATE_DATA
           End If
Application.ScreenUpdating = True
End Sub

Sub CopySheetFromWorkbook()
    Dim sourceWB As Workbook
    Dim currentWB As Workbook: Set currentWB = ThisWorkbook
    Dim sourcePath, sourcebook, userChoice As String
    Dim sourceSheet As Worksheet
    Dim destinationSheet As Worksheet
    Set tbCCY = ThisWorkbook.Sheets("WB_Data").ListObjects("tbCCY")
    sourcePath = tbCCY.ListColumns("Path").DataBodyRange(2, 1).value & "\" & tbCCY.ListColumns("Path").DataBodyRange(1, 1).value & ".csv"
    ' Define the path and workbook name (modify Path1 accordingly)
    sourcebook = tbCCY.ListColumns("Path").DataBodyRange(1, 1).value & ".csv"

    ' Open the source workbook (Read-Only to avoid accidental changes)
    Set sourceWB = Workbooks.Open(sourcePath, ReadOnly:=True)
    
    ' Reference the first (and only) sheet in the source workbook
    Set sourceSheet = sourceWB.Sheets(1)
    
    ' Reference the destination sheet in your current workbook
    Set destinationSheet = currentWB.Sheets("Data Entry")
    
    ' Clear the destination sheet (optional: in case you want to reset it)
    destinationSheet.Cells.Clear
    
    ' Copy the entire sheet's content to "Data" sheet at cell A1
    sourceSheet.UsedRange.Copy destinationSheet.Range("A1")
    
    ' Close the source workbook (without saving)
    sourceWB.Close SaveChanges:=False
    
    ExtractTransactionDataWithDeletion
    
End Sub
Sub CREDIT()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Data Entry")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim transactionMonth, A, PorteM As Integer
    Dim ExtractNumber As Double
    Dim i As Long
    Dim data As Variant
    Dim lineNumber As Long, transactionDate, MyMonth As String
    Dim extractedValue As String, CRY As String
    Dim CodeM As String
    Set tbCCY = ThisWorkbook.Sheets("WB_Data").ListObjects("tbCCY")
    CRY = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
    CodeM = ThisWorkbook.Sheets("Library").Range("R14").value
    PorteM = 0
    
    For i = 1 To lastRow
    If IsDate(ws.Cells(i, "A").value) Then
    transactionDate = Format(CDate(ws.Cells(i, "A")), CodeM & "dd mmm")
    transactionMonth = Month(CDate(ws.Cells(i, "A")))
    ws.Cells(i, "E").value = transactionMonth
    Else
    PorteM = 1
    MyMonth = ws.Cells(i, "A").value
    End If
    If Left(ws.Cells(i, "C").value, 2) = "CR" Then
       extractedValue = Replace(ws.Cells(i, "C").value, "CR", "") ' Remove "CR"
       extractedValue = Replace(extractedValue, "$", "")
       extractedValue = Replace(Replace(extractedValue, " ", ""), Chr(160), "")
       extractedValue = Trim(extractedValue)
       extractedValue = Replace(extractedValue, localDec, altDec)
         If vbaSep = "." Then
            extractedValue = Replace(extractedValue, altDec, Application.International(xlDecimalSeparator))
         End If
         ExtractNumber = CDbl(extractedValue)
         A = 1
    Else
    A = 0
    extractedValue = Replace(ws.Cells(i, "C").value, "$", "")
    extractedValue = Replace(Replace(extractedValue, " ", ""), Chr(160), "")
    extractedValue = Trim(extractedValue)
    extractedValue = Replace(extractedValue, localDec, altDec)
    If vbaSep = "." Then
    extractedValue = Replace(extractedValue, altDec, Application.International(xlDecimalSeparator))
    End If
    ExtractNumber = CDbl(extractedValue)
    End If
    If A = 1 Then
    ws.Cells(i, "C").value = ""
    ws.Cells(i, "D").value = ExtractNumber
    ElseIf A = 0 Then
    ExtractNumber = -ExtractNumber
    ws.Cells(i, "C").value = ExtractNumber
    End If
    If PorteM = 0 Then
    ws.Cells(i, "B").value = "Line " & Format(i, "00000") & " - " & transactionDate & " - " & ws.Cells(i, "B").value & " - (" & ExtractNumber & CRY & ")"
    Else
    ws.Cells(i, "B").value = "Line " & Format(i, "00000") & " - " & MyMonth & " - " & ws.Cells(i, "B").value & " - (" & ExtractNumber & CRY & ")"
    End If
    PorteM = 0
    Next i

End Sub

Sub PARDEFAUT()
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Data Entry")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim transactionMonth, A, C As Integer
    Dim ExtractNumber As Double
    Dim i As Long
    Dim data As Variant
    Dim lineNumber As Long, transactionDate As String
    Dim extractedValue As String, CRY As String
    Dim CodeM As String
    Dim LibR As String

    Set tbCCY = ThisWorkbook.Sheets("WB_Data").ListObjects("tbCCY")
    CRY = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
    CodeM = ThisWorkbook.Sheets("Library").Range("R14").value
    C = 3
    For i = 1 To lastRow
    If IsDate(ws.Cells(i, "A").value) Then
    transactionDate = Format(CDate(ws.Cells(i, "A")), CodeM & "dd mmm")
    transactionMonth = Month(CDate(ws.Cells(i, "A")))
    ws.Cells(i, "E").value = transactionMonth
    End If
    If ws.Cells(i, "C").value <> "" Then
    A = 0
    C = 0
    extractedValue = Replace(ws.Cells(i, "C").value, "$", "")
    extractedValue = Trim(extractedValue)
       extractedValue = Replace(Replace(extractedValue, " ", ""), Chr(160), "")
       extractedValue = Trim(extractedValue)
       extractedValue = Replace(extractedValue, localDec, altDec)
         If vbaSep = "." Then
            extractedValue = Replace(extractedValue, altDec, Application.International(xlDecimalSeparator))
         End If
         ExtractNumber = CDbl(extractedValue)
    ws.Cells(i, "D").value = ""
    Else
    C = 1
    End If
    If ws.Cells(i, "D").value <> "" And C = 1 Then
    A = 1
    extractedValue = Replace(ws.Cells(i, "D").value, "$", "")
    extractedValue = Trim(extractedValue)
    extractedValue = Replace(Replace(extractedValue, " ", ""), Chr(160), "")
       extractedValue = Trim(extractedValue)
       extractedValue = Replace(extractedValue, localDec, altDec)
         If vbaSep = "." Then
            extractedValue = Replace(extractedValue, altDec, Application.International(xlDecimalSeparator))
         End If
         ExtractNumber = CDbl(extractedValue)
    ws.Cells(i, "C").value = ""
    Else
       If C = 0 And ws.Cells(i, "D").value <> "" Then
          LibR = ThisWorkbook.Sheets("Library").Range("A249").value
          MsgBox (LibR), vbOK
          ER = 1
          Exit Sub
       ElseIf C = 1 Then
          LibR = ThisWorkbook.Sheets("Library").Range("A250").value
          MsgBox (LibR), vbOK
          ER = 1
          Exit Sub
       End If
    End If
    C = 3
    If A = 1 Then
    ws.Cells(i, "C").value = ""
    ws.Cells(i, "D").value = ExtractNumber
    ElseIf A = 0 Then
    If ExtractNumber > 0 Then
      ExtractNumber = -ExtractNumber
    End If
    ws.Cells(i, "C").value = ExtractNumber
    End If
    ws.Cells(i, "B").value = "Line " & Format(i, "00000") & " - " & transactionDate & " - " & ws.Cells(i, "B").value & " - (" & ExtractNumber & CRY & ")"
    Next i

End Sub
Sub Reset_Auto()
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Data Entry")
ws.Cells.Clear
End Sub
Sub SG()
Dim sourceWB As Workbook
    Dim currentWB As Workbook: Set currentWB = ThisWorkbook
    Dim sourcePath, sourcebook, userChoice As String
    Dim sourceSheet As Worksheet
    Dim destinationSheet As Worksheet
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Data Entry")
    Dim lastRow As Long
    Dim transactionMonth As Integer
    Dim i As Long, outputRow As Long
    Dim data As Variant
    Dim lineNumber As Long, transactionDate As String, description As String, CRY As String
    Dim value7 As Variant, value8 As Variant, extractedValue As Variant
    Dim skippedCount As Long: skippedCount = 0
    Dim CodeM As String
    Dim delimiter As String
    Dim shapeGroup As Shape

    Set tbCCY = ThisWorkbook.Sheets("WB_Data").ListObjects("tbCCY")
    sourcePath = tbCCY.ListColumns("Path").DataBodyRange(2, 1).value & "\" & tbCCY.ListColumns("Path").DataBodyRange(1, 1).value & ".csv"
    ' Define the path and workbook name (modify Path1 accordingly)
    sourcebook = tbCCY.ListColumns("Path").DataBodyRange(1, 1).value

    Application.DisplayAlerts = False
    ' Open the source workbook (Read-Only to avoid accidental changes)
    Set sourceWB = Workbooks.Open(sourcePath)
    
    delimiter = Application.International(xlListSeparator)
    If delimiter = "," Then
        sourceWB.Sheets(1).Columns(1).TextToColumns _
            Destination:=Range("A1"), DataType:=xlDelimited, Comma:=True
    ' If the delimiter is semicolon, use Semicolon delimiter for splitting data
    ElseIf delimiter = ";" Then
        sourceWB.Sheets(1).Columns(1).TextToColumns _
            Destination:=Range("A1"), DataType:=xlDelimited, Semicolon:=True
    End If
    
    
    
    ' Reference the first (and only) sheet in the source workbook
    Set sourceSheet = sourceWB.Sheets(1)
    
    ' Reference the destination sheet in your current workbook
    Set destinationSheet = currentWB.Sheets("Data Entry")
    
    ' Clear the destination sheet (optional: in case you want to reset it)
    destinationSheet.Cells.Clear
    Application.DisplayAlerts = True
    ' Copy the entire sheet's content to "Data" sheet at cell A1
    sourceSheet.UsedRange.Copy destinationSheet.Range("A1")
    
    ' Close the source workbook (without saving)
    sourceWB.Close SaveChanges:=False

    ws.Rows("1:3").Delete Shift:=xlUp
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ws.Columns("E:E").ClearContents
    ws.Columns("C").Cut
    ws.Columns("B").Insert Shift:=xlToRight
    Application.CutCopyMode = False
    ws.Columns("D").Cut
    ws.Columns("C").Insert Shift:=xlToRight
    Application.CutCopyMode = False
    ws.Columns("D:D").ClearContents
    CodeM = ThisWorkbook.Sheets("Library").Range("R14").value
    CRY = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
    outputRow = 1 ' Initialize output for Column E

           For i = 1 To lastRow
               transactionDate = Format(CDate(Trim(ws.Cells(i, "A").value)), CodeM & "dd mmm") ' Data4: Date (formatted)
               transactionMonth = Month(CDate(Trim(ws.Cells(i, "A").value)))
               description = ws.Cells(i, "B").value ' Data6: Description
               extractedValue = ws.Cells(i, "C").value
        '    Output formatted data in Column B
             ws.Cells(i, "B").value = transactionDate & " - " & description & " / (" & extractedValue & CRY & ")"
        
        ' Handle "Virement Interac" special case
        If extractedValue >= 0 Then
            ws.Cells(i, "D").value = extractedValue ' Revenue in Column C
            ws.Cells(i, "C").value = ""
        End If
        ws.Cells(i, "E").value = transactionMonth

    Next i

    
    ' Loop through each shape in the group and resize it
    On Error Resume Next ' In case there is no shape group
    Set shapeGroup = ws.Shapes("Group 31") ' Replace "GroupName" with the actual name of your shape group
    
    If Not shapeGroup Is Nothing Then
        ' Resize the height of the shape group
        shapeGroup.Height = 190
    End If
    On Error GoTo 0 ' Reset error handling


End Sub
