VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExtraP 
   Caption         =   "Commencer l'extrapolation des données:"
   ClientHeight    =   5620
   ClientLeft      =   112
   ClientTop       =   448
   ClientWidth     =   5360
   OleObjectBlob   =   "ExtraP.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ExtraP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tbline, tbMonth As ListObject
Public Year, A, ER As Integer
Public Monthvalue, BDG As String

Private Sub An_Change()
Dim Mois, Builtarray, LibR1, LibR2 As String
Dim i, YRow As Integer
Dim MonthArray As Variant
Me.An.BorderStyle = fmBorderStyleNone
If Me.An <> "" Then
Set tbMonth = ThisWorkbook.Sheets("WB_Data").ListObjects("Month")
Set tbline = ThisWorkbook.Sheets("WB_Data").ListObjects("Data_Line")
Year = Me.An.value
On Error Resume Next
Set FRow = tbline.ListColumns("Année").DataBodyRange.Find(What:=Year, LookIn:=xlValues, LookAt:=xlWhole)
           
    If Not FRow Is Nothing Then
            YRow = FRow.Row - tbline.DataBodyRange.Row + 1
            Mois = ""
            Builtarray = ""
               mi = 1
                For i = 2 To 13
                        If tbline.ListColumns(i).DataBodyRange(YRow, 1).value <> "" Then
                                Mois = tbMonth.ListColumns("Mois").DataBodyRange(mi, 1).value
                                  If Builtarray = "" Then
                                     Builtarray = Mois
                                  Else
                                     Builtarray = Builtarray & ", " & Mois
                                  End If
  
                         End If
                 mi = mi + 1
                 Next i
  
            MonthArray = Split(Builtarray, ", ")

            Me.Mois.Clear
                    For i = LBound(MonthArray) To UBound(MonthArray)
                            Me.Mois.AddItem MonthArray(i)
                    Next i
    Else
      LibR1 = ThisWorkbook.Sheets("Library").Range("A89").value
      LibR2 = ThisWorkbook.Sheets("Library").Range("A90").value
      MsgBox (LibR1 & vbCrLf & vbCrLf & LibR2), vbOKOnly
    End If
End If
End Sub

Private Sub ComboBdg_Change()
Me.ComboBdg.BorderStyle = fmBorderStyleNone
End Sub

Private Sub CommandButton1_Click()
ER = 0
If Me.An.value = "" Then
   Me.An.BorderStyle = fmBorderStyleSingle
   Me.An.BorderColor = vbRed
   ER = 1
End If

If Me.Mois.value = "" Then
   Me.Mois.BorderStyle = fmBorderStyleSingle
   Me.Mois.BorderColor = vbRed
   ER = 1
End If

    If Me.ComboBdg.value = "" Then
           Me.ComboBdg.BorderStyle = fmBorderStyleSingle
           Me.ComboBdg.BorderColor = vbRed
           ER = 1
    Else
          BDG = Me.ComboBdg.value
    End If
If ER = 1 Then
Exit Sub
End If
Year = Me.An.value
Monthvalue = Me.Mois.value
A = 1
Me.Hide
End Sub

Private Sub Mois_Change()
Me.Mois.BorderStyle = fmBorderStyleNone
End Sub

Private Sub UserForm_Initialize()
Dim LibR As String
SetUserFormBorderColor Me, RGB(0, 0, 0)



LibR = ThisWorkbook.Sheets("Library").Range("A253").value
Me.Label5.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A254").value
Me.Label6.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A255").value
Me.Label7.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A256").value
Me.Label8.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("S35").value
Me.Label9.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("S36").value
Me.Label10.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A83").value
Me.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A84").value
Me.Label1.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A85").value
Me.Label4.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A86").value
Me.Label2.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A87").value
Me.Label3.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A88").value
Me.CommandButton1.Caption = LibR
End Sub
