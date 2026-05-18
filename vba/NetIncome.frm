VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NetIncome 
   Caption         =   "Taxes Braket"
   ClientHeight    =   6230
   ClientLeft      =   112
   ClientTop       =   448
   ClientWidth     =   10376
   OleObjectBlob   =   "NetIncome.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "NetIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public localDec, altDec, vbaSep As String


Private Sub BF1_Change()
Me.BF1.BorderStyle = fmBorderStyleNone
Me.BF1.ForeColor = vbWhite
End Sub

Private Sub BF2_Change()
Me.BF2.BorderStyle = fmBorderStyleNone
Me.BF2.ForeColor = vbWhite
End Sub

Private Sub BF3_Change()
Me.BF3.BorderStyle = fmBorderStyleNone
Me.BF3.ForeColor = vbWhite
End Sub

Private Sub BF4_Change()
Me.BF4.BorderStyle = fmBorderStyleNone
Me.BF4.ForeColor = vbWhite
End Sub

Private Sub BQ1_Change()
Me.BQ1.BorderStyle = fmBorderStyleNone
Me.BQ1.ForeColor = vbWhite
End Sub

Private Sub BQ2_Change()
Me.BQ2.BorderStyle = fmBorderStyleNone
Me.BQ2.ForeColor = vbWhite
End Sub

Private Sub BQ3_Change()
Me.BQ3.BorderStyle = fmBorderStyleNone
Me.BQ3.ForeColor = vbWhite
End Sub

Private Sub BQ4_Change()
Me.BQ4.BorderStyle = fmBorderStyleNone
Me.BQ4.ForeColor = vbWhite
End Sub

Private Sub CommandButton1_Click()
Dim ValRF1, ValRF2, ValRF3, ValRF4, ValRF5, ValRQ1, ValRQ2, ValRQ3, ValRQ4, ValBF1, ValBF2, ValBF3, ValBF4, ValBQ1, ValBQ2, ValBQ3, ValBQ4, ValSemaine As Double
Dim B As Integer
Dim MyVal As String
Dim VVal As Variant
B = 0

If Me.RF1.value = "" Then
    Me.RF1.BorderStyle = fmBorderStyleSingle
    Me.RF1.BorderColor = vbRed
    B = 1
    Else
    VVal = Replace(Me.RF1.value, localDec, altDec)
    If vbaSep = "." Then
    VVal = Replace(VVal, altDec, Application.International(xlDecimalSeparator))
    End If
    VVal = Trim(VVal)
    If IsNumeric(VVal) Then
    ValRF1 = CDbl(VVal)
    Me.RF1.value = ValRF1
    Else
    Me.RF1.ForeColor = vbRed
    B = 1
    End If
    End If
    
If Me.RF2.value = "" Then
    Me.RF2.BorderStyle = fmBorderStyleSingle
    Me.RF2.BorderColor = vbRed
    B = 1
    Else
    VVal = Replace(Me.RF2.value, localDec, altDec)
    If vbaSep = "." Then
    VVal = Replace(VVal, altDec, Application.International(xlDecimalSeparator))
    End If
    VVal = Trim(VVal)
    If IsNumeric(VVal) Then
    ValRF2 = CDbl(VVal)
    Me.RF2.value = ValRF2
    Else
    Me.RF2.ForeColor = vbRed
    B = 1
    End If
    End If
    
If Me.RF3.value = "" Then
    Me.RF3.BorderStyle = fmBorderStyleSingle
    Me.RF3.BorderColor = vbRed
    B = 1
    Else
    VVal = Replace(Me.RF3.value, localDec, altDec)
    If vbaSep = "." Then
    VVal = Replace(VVal, altDec, Application.International(xlDecimalSeparator))
    End If
    VVal = Trim(VVal)
    If IsNumeric(VVal) Then
    ValRF3 = CDbl(VVal)
    Me.RF3.value = ValRF3
    Else
    Me.RF3.ForeColor = vbRed
    B = 1
    End If
    End If
    
If Me.RF4.value = "" Then
    Me.RF4.BorderStyle = fmBorderStyleSingle
    Me.RF4.BorderColor = vbRed
    B = 1
    Else
    VVal = Replace(Me.RF4.value, localDec, altDec)
    If vbaSep = "." Then
    VVal = Replace(VVal, altDec, Application.International(xlDecimalSeparator))
    End If
    VVal = Trim(VVal)
    If IsNumeric(VVal) Then
    ValRF4 = CDbl(VVal)
    Me.RF4.value = ValRF4
    Else
    Me.RF4.ForeColor = vbRed
    B = 1
    End If
    End If
    
If Me.RF5.value = "" Then
    Me.RF5.BorderStyle = fmBorderStyleSingle
    Me.RF5.BorderColor = vbRed
    B = 1
    Else
    VVal = Replace(Me.RF5.value, localDec, altDec)
    If vbaSep = "." Then
    VVal = Replace(VVal, altDec, Application.International(xlDecimalSeparator))
    End If
    VVal = Trim(VVal)
    If IsNumeric(VVal) Then
    ValRF5 = CDbl(VVal)
    Me.RF5.value = ValRF5
    Else
    Me.RF5.ForeColor = vbRed
    B = 1
    End If
    End If

If Me.RQ1.value = "" Then
    Me.RQ1.BorderStyle = fmBorderStyleSingle
    Me.RQ1.BorderColor = vbRed
    B = 1
    Else
    VVal = Replace(Me.RQ1.value, localDec, altDec)
    If vbaSep = "." Then
    VVal = Replace(VVal, altDec, Application.International(xlDecimalSeparator))
    End If
    VVal = Trim(VVal)
    If IsNumeric(VVal) Then
    ValRQ1 = CDbl(VVal)
    Me.RQ1.value = ValRQ1
    Else
    Me.RQ1.ForeColor = vbRed
    B = 1
    End If
    End If
    
If Me.RQ2.value = "" Then
    Me.RQ2.BorderStyle = fmBorderStyleSingle
    Me.RQ2.BorderColor = vbRed
    B = 1
    Else
    VVal = Replace(Me.RQ2.value, localDec, altDec)
    If vbaSep = "." Then
    VVal = Replace(VVal, altDec, Application.International(xlDecimalSeparator))
    End If
    VVal = Trim(VVal)
    If IsNumeric(VVal) Then
    ValRQ2 = CDbl(VVal)
    Me.RQ2.value = ValRQ2
    Else
    Me.RQ2.ForeColor = vbRed
    B = 1
    End If
    End If
    
If Me.RQ3.value = "" Then
    Me.RQ3.BorderStyle = fmBorderStyleSingle
    Me.RQ3.BorderColor = vbRed
    B = 1
    Else
    VVal = Replace(Me.RQ3.value, localDec, altDec)
    If vbaSep = "." Then
    VVal = Replace(VVal, altDec, Application.International(xlDecimalSeparator))
    End If
    VVal = Trim(VVal)
    If IsNumeric(VVal) Then
    ValRQ3 = CDbl(VVal)
    Me.RQ3.value = ValRQ3
    Else
    Me.RQ3.ForeColor = vbRed
    B = 1
    End If
    End If
    
If Me.RQ4.value = "" Then
    Me.RQ4.BorderStyle = fmBorderStyleSingle
    Me.RQ4.BorderColor = vbRed
    B = 1
    Else
    VVal = Replace(Me.RQ4.value, localDec, altDec)
    If vbaSep = "." Then
    VVal = Replace(VVal, altDec, Application.International(xlDecimalSeparator))
    End If
    VVal = Trim(VVal)
    If IsNumeric(VVal) Then
    ValRQ4 = CDbl(VVal)
    Me.RQ4.value = ValRQ4
    Else
    Me.RQ4.ForeColor = vbRed
    B = 1
    End If
    End If

If Me.BF1.value = "" Then
    Me.BF1.BorderStyle = fmBorderStyleSingle
    Me.BF1.BorderColor = vbRed
    B = 1
    Else
    VVal = Replace(Me.BF1.value, localDec, altDec)
    If vbaSep = "." Then
    VVal = Replace(VVal, altDec, Application.International(xlDecimalSeparator))
    End If
    VVal = Trim(VVal)
    If IsNumeric(VVal) Then
    ValBF1 = CDbl(VVal)
    Me.BF1.value = ValBF1
    Else
    Me.BF1.ForeColor = vbRed
    B = 1
    End If
    End If
    
If Me.BF2.value = "" Then
    Me.BF2.BorderStyle = fmBorderStyleSingle
    Me.BF2.BorderColor = vbRed
    B = 1
    Else
    VVal = Replace(Me.BF2.value, localDec, altDec)
    If vbaSep = "." Then
    VVal = Replace(VVal, altDec, Application.International(xlDecimalSeparator))
    End If
    VVal = Trim(VVal)
    If IsNumeric(VVal) Then
    ValBF2 = CDbl(VVal)
    Me.BF2.value = ValBF2
    Else
    Me.BF2.ForeColor = vbRed
    B = 1
    End If
    End If
    
If Me.BF3.value = "" Then
    Me.BF3.BorderStyle = fmBorderStyleSingle
    Me.BF3.BorderColor = vbRed
    B = 1
    Else
    VVal = Replace(Me.BF3.value, localDec, altDec)
    If vbaSep = "." Then
    VVal = Replace(VVal, altDec, Application.International(xlDecimalSeparator))
    End If
    VVal = Trim(VVal)
    If IsNumeric(VVal) Then
    ValBF3 = CDbl(VVal)
    Me.BF3.value = ValBF3
    Else
    Me.BF3.ForeColor = vbRed
    B = 1
    End If
    End If
    
If Me.BF4.value = "" Then
    Me.BF4.BorderStyle = fmBorderStyleSingle
    Me.BF4.BorderColor = vbRed
    B = 1
    Else
    VVal = Replace(Me.BF4.value, localDec, altDec)
    If vbaSep = "." Then
    VVal = Replace(VVal, altDec, Application.International(xlDecimalSeparator))
    End If
    VVal = Trim(VVal)
    If IsNumeric(VVal) Then
    ValBF4 = CDbl(VVal)
    Me.BF4.value = ValBF4
    Else
    Me.BF4.ForeColor = vbRed
    B = 1
    End If
    End If

If Me.BQ1.value = "" Then
    Me.BQ1.BorderStyle = fmBorderStyleSingle
    Me.BQ1.BorderColor = vbRed
    B = 1
    Else
    VVal = Replace(Me.BQ1.value, localDec, altDec)
    If vbaSep = "." Then
    VVal = Replace(VVal, altDec, Application.International(xlDecimalSeparator))
    End If
    VVal = Trim(VVal)
    If IsNumeric(VVal) Then
    ValBQ1 = CDbl(VVal)
    Me.BQ1.value = ValBQ1
    Else
    Me.BQ1.ForeColor = vbRed
    B = 1
    End If
    End If
    
If Me.BQ2.value = "" Then
    Me.BQ2.BorderStyle = fmBorderStyleSingle
    Me.BQ2.BorderColor = vbRed
    B = 1
    Else
    VVal = Replace(Me.BQ2.value, localDec, altDec)
    If vbaSep = "." Then
    VVal = Replace(VVal, altDec, Application.International(xlDecimalSeparator))
    End If
    VVal = Trim(VVal)
    If IsNumeric(VVal) Then
    ValBQ2 = CDbl(VVal)
    Me.BQ2.value = ValBQ2
    Else
    Me.BQ2.ForeColor = vbRed
    B = 1
    End If
    End If
    
If Me.BQ3.value = "" Then
    Me.BQ3.BorderStyle = fmBorderStyleSingle
    Me.BQ3.BorderColor = vbRed
    B = 1
    Else
    VVal = Replace(Me.BQ3.value, localDec, altDec)
    If vbaSep = "." Then
    VVal = Replace(VVal, altDec, Application.International(xlDecimalSeparator))
    End If
    VVal = Trim(VVal)
    If IsNumeric(VVal) Then
    ValBQ3 = CDbl(VVal)
    Me.BQ3.value = ValBQ3
    Else
    Me.BQ3.ForeColor = vbRed
    B = 1
    End If
    End If
    
If Me.BQ4.value = "" Then
    Me.BQ4.BorderStyle = fmBorderStyleSingle
    Me.BQ4.BorderColor = vbRed
    B = 1
    Else
    VVal = Replace(Me.BQ4.value, localDec, altDec)
    If vbaSep = "." Then
    VVal = Replace(VVal, altDec, Application.International(xlDecimalSeparator))
    End If
    VVal = Trim(VVal)
    If IsNumeric(VVal) Then
    ValBQ4 = CDbl(VVal)
    Me.BQ4.value = ValBQ4
    Else
    Me.BQ4.ForeColor = vbRed
    B = 1
    End If
    End If

If Me.Semaine.value = "" Then
    Me.Semaine.BorderStyle = fmBorderStyleSingle
    Me.Semaine.BorderColor = vbRed
    B = 1
    Else
    VVal = Replace(Me.Semaine.value, localDec, altDec)
    If vbaSep = "." Then
    VVal = Replace(VVal, altDec, Application.International(xlDecimalSeparator))
    End If
    VVal = Trim(VVal)
    If IsNumeric(VVal) Then
    ValSemaine = CDbl(VVal)
    Me.Semaine.value = ValSemaine
    Else
    Me.Semaine.ForeColor = vbRed
    B = 1
    End If
    End If

If B = 1 Then
Exit Sub
End If

ThisWorkbook.Sheets("Sheet1").Range("B33").value = ValRF1
ThisWorkbook.Sheets("Sheet1").Range("B34").value = ValRF2
ThisWorkbook.Sheets("Sheet1").Range("B35").value = ValRF3
ThisWorkbook.Sheets("Sheet1").Range("B36").value = ValRF4
ThisWorkbook.Sheets("Sheet1").Range("B37").value = ValRF5

ThisWorkbook.Sheets("Sheet1").Range("B38").value = ValRQ1
ThisWorkbook.Sheets("Sheet1").Range("B39").value = ValRQ2
ThisWorkbook.Sheets("Sheet1").Range("B40").value = ValRQ3
ThisWorkbook.Sheets("Sheet1").Range("B41").value = ValRQ4

ThisWorkbook.Sheets("Sheet1").Range("C33").value = ValBF1
ThisWorkbook.Sheets("Sheet1").Range("C34").value = ValBF2
ThisWorkbook.Sheets("Sheet1").Range("C35").value = ValBF3
ThisWorkbook.Sheets("Sheet1").Range("C36").value = ValBF4

ThisWorkbook.Sheets("Sheet1").Range("C38").value = ValBQ1
ThisWorkbook.Sheets("Sheet1").Range("C39").value = ValBQ2
ThisWorkbook.Sheets("Sheet1").Range("C40").value = ValBQ3
ThisWorkbook.Sheets("Sheet1").Range("C41").value = ValBQ4

ThisWorkbook.Sheets("Sheet1").Range("B2").value = ValSemaine

Unload Me
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub Label28_Click()
Dim url As String
    url = "https://www.canada.ca/en/revenue-agency/services/tax/individuals/frequently-asked-questions-individuals/canadian-income-tax-rates-individuals-current-previous-years.html"

    ThisWorkbook.FollowHyperlink Address:=url, NewWindow:=True
End Sub

Private Sub Label29_Click()
Dim url As String
    url = "https://www.revenuquebec.ca/en/citizens/your-situation/new-residents/the-quebec-tax-system/income-tax-rates/"

    ThisWorkbook.FollowHyperlink Address:=url, NewWindow:=True

End Sub


Private Sub RF1_Change()
Me.RF1.BorderStyle = fmBorderStyleNone
Me.RF1.ForeColor = vbWhite
End Sub

Private Sub RF2_Change()
Me.RF2.BorderStyle = fmBorderStyleNone
Me.RF2.ForeColor = vbWhite
End Sub

Private Sub RF3_Change()
Me.RF3.BorderStyle = fmBorderStyleNone
Me.RF3.ForeColor = vbWhite
End Sub

Private Sub RF4_Change()
Me.RF4.BorderStyle = fmBorderStyleNone
Me.RF4.ForeColor = vbWhite
End Sub

Private Sub RF5_Change()
Me.RF5.BorderStyle = fmBorderStyleNone
Me.RF5.ForeColor = vbWhite
End Sub

Private Sub RQ1_Change()
Me.RQ1.BorderStyle = fmBorderStyleNone
Me.RQ1.ForeColor = vbWhite
End Sub

Private Sub RQ2_Change()
Me.RQ2.BorderStyle = fmBorderStyleNone
Me.RQ2.ForeColor = vbWhite
End Sub

Private Sub RQ3_Change()
Me.RQ3.BorderStyle = fmBorderStyleNone
Me.RQ3.ForeColor = vbWhite
End Sub

Private Sub RQ4_Change()
Me.RQ4.BorderStyle = fmBorderStyleNone
Me.RQ4.ForeColor = vbWhite
End Sub

Private Sub Semaine_Change()
Me.Semaine.BorderStyle = fmBorderStyleNone
Me.Semaine.ForeColor = vbWhite
End Sub

Private Sub UserForm_Initialize()
Dim LibR As String
LibR = ThisWorkbook.Sheets("Library").Range("A130").value
Me.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A131").value
Me.Label11.Caption = LibR
Me.Label25.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A132").value
Me.Label10.Caption = LibR
Me.Label24.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A133").value
Me.Label9.Caption = LibR
Me.Label23.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A134").value
Me.Label8.Caption = LibR
Me.Label22.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A135").value
Me.Label14.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A136").value
Me.Label13.Caption = LibR
Me.Label27.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A137").value
Me.Label12.Caption = LibR
Me.Label26.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A138").value
Me.Label30.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A139").value
Me.Label29.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A140").value
Me.Label28.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A141").value
Me.CommandButton1.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A142").value
Me.CommandButton2.Caption = LibR

 localDec = Application.International(xlDecimalSeparator)
    
    ' Define the alternative decimal separator: if local is dot then alt is comma; otherwise, alt is dot.
    If localDec = "." Then
        altDec = ","
    Else
        altDec = "."
    End If

vbaSep = Mid$(Format(1.1, "0.0"), 2, 1)

Me.Semaine.value = ThisWorkbook.Sheets("Sheet1").Range("B2").value

Me.RF1.value = ThisWorkbook.Sheets("Sheet1").Range("B33").value
Me.RF2.value = ThisWorkbook.Sheets("Sheet1").Range("B34").value
Me.RF3.value = ThisWorkbook.Sheets("Sheet1").Range("B35").value
Me.RF4.value = ThisWorkbook.Sheets("Sheet1").Range("B36").value
Me.RF5.value = ThisWorkbook.Sheets("Sheet1").Range("B37").value

Me.RQ1.value = ThisWorkbook.Sheets("Sheet1").Range("B38").value
Me.RQ2.value = ThisWorkbook.Sheets("Sheet1").Range("B39").value
Me.RQ3.value = ThisWorkbook.Sheets("Sheet1").Range("B40").value
Me.RQ4.value = ThisWorkbook.Sheets("Sheet1").Range("B41").value

Me.BF1.value = ThisWorkbook.Sheets("Sheet1").Range("C33").value
Me.BF2.value = ThisWorkbook.Sheets("Sheet1").Range("C34").value
Me.BF3.value = ThisWorkbook.Sheets("Sheet1").Range("C35").value
Me.BF4.value = ThisWorkbook.Sheets("Sheet1").Range("C36").value

Me.BQ1.value = ThisWorkbook.Sheets("Sheet1").Range("C38").value
Me.BQ2.value = ThisWorkbook.Sheets("Sheet1").Range("C39").value
Me.BQ3.value = ThisWorkbook.Sheets("Sheet1").Range("C40").value
Me.BQ4.value = ThisWorkbook.Sheets("Sheet1").Range("C41").value

End Sub
