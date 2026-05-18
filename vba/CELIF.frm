VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CELIF 
   Caption         =   "Plafond CELI"
   ClientHeight    =   4110
   ClientLeft      =   112
   ClientTop       =   448
   ClientWidth     =   5632
   OleObjectBlob   =   "CELIF.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CELIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tbCELY, tbCCY As ListObject
Public B As Integer
Public localDec, altDec, MyPlafond, vbaSep As String
Public PLF As Double
Public VPlafond As Variant
Private Sub CELI_Click()
Dim url As String
    url = "" ' Official Registred_account contribution information URL

    ThisWorkbook.FollowHyperlink Address:=url, NewWindow:=True
End Sub

Private Sub CommandButton1_Click()
Me.ErrorY.ForeColor = &H808080
B = 0
If Me.Plafond.value <> "" Then
    VPlafond = Replace(Me.Plafond.value, localDec, altDec)
    If vbaSep = "." Then
    VPlafond = Replace(VPlafond, altDec, Application.International(xlDecimalSeparator))
    End If
    VPlafond = Trim(VPlafond)
    If IsNumeric(VPlafond) Then
       PLF = CDbl(VPlafond)
        i = tbCELY.ListColumns("First Year").DataBodyRange(8, 1).value
        Me.Plafond.value = PLF
        tbCELY.ListColumns("Plafond").DataBodyRange(i, 1).value = PLF
        B = 1
    Else
    Me.Plafond.ForeColor = vbRed
    Exit Sub
    End If
Else
Me.Plafond.BorderStyle = fmBorderStyleSingle
Me.Plafond.BorderColor = vbRed
End If
If Me.Ann�e.value <> "" Then
      If IsNumeric(Me.Ann�e.value) Then
            If Abs(Me.Ann�e.value) >= 1000 And Abs(Me.Ann�e.value) <= 9999 Then
            tbCELY.ListColumns("First Year").DataBodyRange(1, 1).value = Me.Ann�e.value
            B = 1
            Else
            Me.ErrorY.ForeColor = vbRed
            Me.Ann�e.value = ""
            Exit Sub
            End If
      Else
            Me.ErrorY.ForeColor = vbRed
            Me.Ann�e.value = ""
            Exit Sub
      End If
Else
Me.Ann�e.value = ""
Me.ErrorY.ForeColor = vbRed
Exit Sub
End If

If B = 1 Then
Unload Me
End If

End Sub



Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub Plafond_Change()
Me.Plafond.BorderStyle = fmBorderStyleNone
Me.Plafond.ForeColor = vbWhite
End Sub

Private Sub Year_Change()
tbCELY.ListColumns("First Year").DataBodyRange(6, 1).value = Me.Year.value
Me.Plafond.value = tbCELY.ListColumns("First Year").DataBodyRange(7, 1).value
End Sub
Private Sub UserForm_Initialize()
Me.ErrorY.ForeColor = &H808080
Dim LibR As String

localDec = Application.International(xlDecimalSeparator)
    
    ' Define the alternative decimal separator: if local is dot then alt is comma; otherwise, alt is dot.
    If localDec = "." Then
        altDec = ","
    Else
        altDec = "."
    End If

vbaSep = Mid$(Format(1.1, "0.0"), 2, 1)

LibR = ThisWorkbook.Sheets("Library").Range("A34").value
Me.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A35").value
Me.Label1.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A36").value
Me.Label2.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A37").value
Me.Label5.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A38").value
Me.CommandButton1.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A41").value
Me.CommandButton2.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A33").value
Me.CELI.Caption = LibR
Set tbCCY = ThisWorkbook.Sheets("WB_Data").ListObjects("tbCCY")
Me.SIG.Caption = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
Set tbCELY = ThisWorkbook.Sheets("Sheet1").ListObjects("Registred_account")
Me.Ann�e.value = tbCELY.ListColumns("First Year").DataBodyRange(1, 1).value


End Sub
