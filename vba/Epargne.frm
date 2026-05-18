VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Epargne 
   Caption         =   "Setting Epargne Cible"
   ClientHeight    =   3130
   ClientLeft      =   112
   ClientTop       =   448
   ClientWidth     =   8656.001
   OleObjectBlob   =   "Epargne.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Epargne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tbline, tbh, tbCCY As ListObject
Public Année, CYear As Integer
Public VCompte, VCible As Variant
Public FRow As Range
Public YRow As Long
Public localDec, altDec, MyCompte, MyCible, vbaSep As String
Private Sub Cible_Change()
Me.Cible.ForeColor = vbWhite
Me.Compte.BorderStyle = fmBorderStyleNone
Me.Cible.BorderStyle = fmBorderStyleNone

End Sub

Private Sub CommandButton1_Click()
Dim ValE, ValC As Double
Dim A, B As Integer
Dim LibR As String
A = 0
B = 0
If Me.Compte.value = "" Or Me.Cible.value = "" Then
Me.Compte.BorderStyle = fmBorderStyleSingle
Me.Compte.BorderColor = vbRed
Me.Cible.BorderStyle = fmBorderStyleSingle
Me.Cible.BorderColor = vbRed
Else
    
    If Me.ListY.value = "" Then
    Me.LabelY.BorderStyle = fmBorderStyleSingle
    Me.LabelY.BorderColor = vbRed
    Exit Sub
    Else
    CYear = Me.ListY.value
    End If
    VCompte = Replace(Me.Compte.value, localDec, altDec)
    If vbaSep = "." Then
    VCompte = Replace(VCompte, altDec, Application.International(xlDecimalSeparator))
    End If
    VCompte = Trim(VCompte)
    If IsNumeric(VCompte) Then
       ValE = CDbl(VCompte)
       A = 1
    Else
    Me.Compte.ForeColor = vbRed
    Exit Sub
    End If
    VCible = Replace(Me.Cible.value, localDec, altDec)
    If vbaSep = "." Then
    VCible = Replace(VCible, altDec, Application.International(xlDecimalSeparator))
    End If
    VCible = Trim(VCible)
    If IsNumeric(VCible) Then
       ValC = CDbl(VCible)
       B = 1
    Else
    Me.Cible.ForeColor = vbRed
    Exit Sub
    End If
    
    On Error Resume Next
       Set FRow = tbline.ListColumns(Année).DataBodyRange.Find(What:=CYear, LookIn:=xlValues, LookAt:=xlWhole)
           If Not FRow Is Nothing Then
                  YRow = FRow.Row - 1
           Else
                  LibR = ThisWorkbook.Sheets("Library").Range("A74").value
                  MsgBox (LibR & CYear & "."), vbOKCancel
                  Exit Sub
           End If
    
       Application.ScreenUpdating = False
       ThisWorkbook.Sheets("WB_Data").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
        If A = 1 And B = 1 Then
          tbline.ListColumns("Solde").DataBodyRange(YRow, 1).value = ValE
        End If
        If B = 1 Then
          tbline.ListColumns("Cible").DataBodyRange(YRow, 1).value = ValC
        End If
        ThisWorkbook.Sheets("WB_Data").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
        ThisWorkbook.Sheets("DashBoard").Activate
        Application.ScreenUpdating = True
        Unload Me
End If

End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub Compte_Change()
Me.Compte.ForeColor = vbWhite
Me.Compte.BorderStyle = fmBorderStyleNone
Me.Cible.BorderStyle = fmBorderStyleNone
End Sub
Private Sub ListY_Change()

CYear = Me.ListY.value
Me.LabelY.BorderStyle = fmBorderStyleNone
On Error Resume Next
       Set FRow = tbline.ListColumns(Année).DataBodyRange.Find(What:=CYear, LookIn:=xlValues, LookAt:=xlWhole)
           If Not FRow Is Nothing Then
                  YRow = FRow.Row - 1
                  Me.Cible.value = tbline.ListColumns("Cible").DataBodyRange(YRow, 1).value
                  Me.Compte.value = tbline.ListColumns("Solde").DataBodyRange(YRow, 1).value
           Else
           Me.Cible.value = ""
           Me.Compte.value = ""
           End If


End Sub


Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
Dim LibR As String

localDec = Application.International(xlDecimalSeparator)
    
    ' Define the alternative decimal separator: if local is dot then alt is comma; otherwise, alt is dot.
    If localDec = "." Then
        altDec = ","
    Else
        altDec = "."
    End If

vbaSep = Mid$(Format(1.1, "0.0"), 2, 1)

LibR = ThisWorkbook.Sheets("Library").Range("A69").value
Me.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A70").value
Me.LabelY.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A71").value
Me.Label1.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A72").value
Me.Label2.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A73").value
Me.CommandButton1.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A79").value
Me.CommandButton2.Caption = LibR

Set tbCCY = ThisWorkbook.Sheets("WB_Data").ListObjects("tbCCY")
Me.SIG1.Caption = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
Me.SIG2.Caption = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
Set tbline = ThisWorkbook.Sheets("WB_Data").ListObjects("Data_Line")
Set tbh = ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead")
Année = tbh.ListColumns("Header1").DataBodyRange(1, 1).value
CYear = ThisWorkbook.Sheets("ControlChart").Range("CM1").value
On Error Resume Next
       Set FRow = tbline.ListColumns(Année).DataBodyRange.Find(What:=CYear, LookIn:=xlValues, LookAt:=xlWhole)
           If Not FRow Is Nothing Then
                  YRow = FRow.Row - 1
           End If
Me.ListY.value = CYear
Me.Cible.value = tbline.ListColumns("Cible").DataBodyRange(YRow, 1).value
Me.Compte.value = tbline.ListColumns("Solde").DataBodyRange(YRow, 1).value
End Sub
