VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EPProgress 
   Caption         =   "Enregistrer votre progression"
   ClientHeight    =   5350
   ClientLeft      =   112
   ClientTop       =   448
   ClientWidth     =   6352
   OleObjectBlob   =   "EPProgress.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "EPProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DTEP, tbCCY As ListObject
Public Year As Integer
Public FRow As Range
Public YRow As Long
Public CAP As Double
Public CRY, MyCapital As String
Public localDec, altDec, vbaSep As String
Public VCapital As Variant


Private Sub Capital_Change()
Me.Capital.BorderStyle = fmBorderStyleNone
Me.Capital.ForeColor = vbWhite
End Sub

Private Sub Commandquitt_Click()
Unload Me
End Sub


Private Sub Frame1_Click()

End Sub

Private Sub ListAN_Click()
Dim LibR As String
If A = 0 Then
   Year = Me.ListAN.value
End If
Set FRow = DTEP.ListColumns("Année").DataBodyRange.Find(What:=Year, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                    On Error Resume Next
                        If Not FRow Is Nothing Then
                           YRow = FRow.Row - 1
                        Else
                        LibR = ThisWorkbook.Sheets("Library").Range("A80").value
                        MsgBox (LibR), vbOK
                        Exit Sub
                        End If
           If DTEP.ListColumns("TTEP").DataBodyRange(YRow, 1).value = "" Then
           Me.Capital.value = 0
           Else
           Me.Capital.value = DTEP.ListColumns("TTEP").DataBodyRange(YRow, 1).value
           End If
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

LibR = ThisWorkbook.Sheets("Library").Range("A76").value
Me.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A77").value
Me.Label2.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A78").value
Me.Valider.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A79").value
Me.Commandquitt.Caption = LibR
A = 1
Set DTEP = ThisWorkbook.Sheets("ControlChart").ListObjects("DTEP")
Year = VBA.DateTime.Year(Date)
Me.ListAN.value = Year
A = 0
Set tbCCY = ThisWorkbook.Sheets("WB_Data").ListObjects("tbCCY")
Me.SIG.Caption = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
CRY = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
End Sub

Private Sub Valider_Click()
Dim LibR As String
If Me.Capital.value = "" Then
    Me.Capital.BorderStyle = fmBorderStyleSingle
    Me.Capital.BorderColor = vbRed
    Exit Sub
    Else
    VCapital = Replace(Me.Capital.value, localDec, altDec)
    If vbaSep = "." Then
    VCapital = Replace(VCapital, altDec, Application.International(xlDecimalSeparator))
    End If
    VCapital = Trim(VCapital)
    If IsNumeric(VCapital) Then
       CAP = CDbl(VCapital)
    Else
    Me.Capital.ForeColor = vbRed
    Exit Sub
    End If
    End If
    If Year = 0 Then
    LibR = ThisWorkbook.Sheets("Library").Range("A81").value
    MsgBox (LibR), vbOK
    End If
    DTEP.ListColumns("TTEP").DataBodyRange(YRow, 1).value = CAP
    DTEP.ListColumns("TTEP").DataBodyRange.NumberFormat = "# ### ###" & CRY
    Me.Capital.BorderStyle = fmBorderStyleSingle
    Me.Capital.BorderColor = vbGreen
End Sub
