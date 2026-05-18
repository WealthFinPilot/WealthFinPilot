VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public BDG As String

Private Sub B1MED_Click()
Dim Papillon As ChartObject
Set Papillon = ThisWorkbook.Sheets("DashBoard").ChartObjects("Papillon1")
 If B1MED.value = True Then
        B1MEP.value = False
        B1LMP.value = False
        B1LMD.value = False
        ThisWorkbook.Sheets("ControlChart_Budget1").Range("DO19").value = 3
        Papillon.Chart.Refresh
    End If
End Sub
Private Sub B1MEP_Click()
Dim Papillon As ChartObject
Set Papillon = ThisWorkbook.Sheets("DashBoard").ChartObjects("Papillon1")
 If B1MEP.value = True Then
        B1MED.value = False
        B1LMP.value = False
        B1LMD.value = False
        ThisWorkbook.Sheets("ControlChart_Budget1").Range("DO19").value = 1
        Papillon.Chart.Refresh
    End If

End Sub
Private Sub B1LMP_Click()
Dim Papillon As ChartObject
Set Papillon = ThisWorkbook.Sheets("DashBoard").ChartObjects("Papillon1")
If B1LMP.value = True Then
        B1MED.value = False
        B1MEP.value = False
        B1LMD.value = False
        ThisWorkbook.Sheets("ControlChart_Budget1").Range("DO19").value = 2
        Papillon.Chart.Refresh
    End If

End Sub
Private Sub B1LMD_Click()
Dim Papillon As ChartObject
Set Papillon = ThisWorkbook.Sheets("DashBoard").ChartObjects("Papillon1")
 If B1LMD.value = True Then
        B1MED.value = False
        B1MEP.value = False
        B1LMP.value = False
        ThisWorkbook.Sheets("ControlChart_Budget1").Range("DO19").value = 4
        Papillon.Chart.Refresh
    End If

End Sub

Private Sub B2MED_Click()
Dim Papillon As ChartObject
Set Papillon = ThisWorkbook.Sheets("DashBoard").ChartObjects("Papillon2")
 If B2MED.value = True Then
        B2MEP.value = False
        B2LMP.value = False
        B2LMD.value = False
        ThisWorkbook.Sheets("ControlChart_Budget2").Range("DO19").value = 3
        Papillon.Chart.Refresh
    End If

End Sub
Private Sub B2MEP_Click()
 If B2MEP.value = True Then
 Dim Papillon As ChartObject
Set Papillon = ThisWorkbook.Sheets("DashBoard").ChartObjects("Papillon2")
        B2MED.value = False
        B2LMP.value = False
        B2LMD.value = False
        ThisWorkbook.Sheets("ControlChart_Budget2").Range("DO19").value = 1
        Papillon.Chart.Refresh
    End If

End Sub
Private Sub B2LMP_Click()
Dim Papillon As ChartObject
Set Papillon = ThisWorkbook.Sheets("DashBoard").ChartObjects("Papillon2")
 If B2LMP.value = True Then
        B2MED.value = False
        B2MEP.value = False
        B2LMD.value = False
        ThisWorkbook.Sheets("ControlChart_Budget2").Range("DO19").value = 2
        Papillon.Chart.Refresh
    End If

End Sub
Private Sub B2LMD_Click()
Dim Papillon As ChartObject
Set Papillon = ThisWorkbook.Sheets("DashBoard").ChartObjects("Papillon2")
 If B2LMD.value = True Then
        B2MED.value = False
        B2MEP.value = False
        B2LMP.value = False
        ThisWorkbook.Sheets("ControlChart_Budget2").Range("DO19").value = 4
        Papillon.Chart.Refresh
    End If

End Sub

Private Sub B2Q1_Click()
Dim GraphBudget As ChartObject
 If B2Q1.value = True Then
        B2Q2.value = False
        B2Q3.value = False
        B2Q4.value = False
        B2Q5.value = False
        ThisWorkbook.Sheets("ControlChart_Budget2").Range("AT4").value = 1
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2AL")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2CH")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2LO")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2VO")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2NE")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2SE")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2AB")
        GraphBudget.Chart.Refresh
    End If
End Sub
Private Sub B2Q2_Click()
Dim GraphBudget As ChartObject
 If B2Q2.value = True Then
        B2Q1.value = False
        B2Q3.value = False
        B2Q4.value = False
        B2Q5.value = False
        ThisWorkbook.Sheets("ControlChart_Budget2").Range("AT4").value = 2
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2AL")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2CH")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2LO")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2VO")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2NE")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2SE")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2AB")
        GraphBudget.Chart.Refresh
    End If
End Sub
Private Sub B2Q3_Click()
Dim GraphBudget As ChartObject
 If B2Q3.value = True Then
        B2Q1.value = False
        B2Q2.value = False
        B2Q4.value = False
        B2Q5.value = False
        ThisWorkbook.Sheets("ControlChart_Budget2").Range("AT4").value = 3
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2AL")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2CH")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2LO")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2VO")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2NE")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2SE")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2AB")
        GraphBudget.Chart.Refresh
    End If
End Sub
Private Sub B2Q4_Click()
Dim GraphBudget As ChartObject
 If B2Q4.value = True Then
        B2Q1.value = False
        B2Q2.value = False
        B2Q3.value = False
        B2Q5.value = False
        ThisWorkbook.Sheets("ControlChart_Budget2").Range("AT4").value = 4
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2AL")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2CH")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2LO")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2VO")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2NE")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2SE")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2AB")
        GraphBudget.Chart.Refresh
    End If
End Sub
Private Sub B2Q5_Click()
Dim GraphBudget As ChartObject
 If B2Q5.value = True Then
        B2Q1.value = False
        B2Q2.value = False
        B2Q3.value = False
        B2Q4.value = False
        ThisWorkbook.Sheets("ControlChart_Budget2").Range("AT4").value = 5
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2AL")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2CH")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2LO")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2VO")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2NE")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2SE")
        GraphBudget.Chart.Refresh
        Set GraphBudget = ThisWorkbook.Sheets("DashBoard").ChartObjects("B2AB")
        GraphBudget.Chart.Refresh
    End If
End Sub

Private Sub CommandButton10_Click()
NetIncome.Show
End Sub

Private Sub CommandButton13_Click()
CELIF.Show
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub Ratio_Click()
Application.ScreenUpdating = False
 If Ratio.value Then
        Cost.value = False
    End If
ThisWorkbook.Sheets("ControlChart").Range("J19").value = 1

Application.ScreenUpdating = True
End Sub
Private Sub Cost_Click()
Application.ScreenUpdating = False

 If Cost.value Then
        Ratio.value = False
    End If
ThisWorkbook.Sheets("ControlChart").Range("J19").value = 2
Application.ScreenUpdating = True
End Sub

Sub Hide()
Dim ws As Worksheet

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("DashBoard") ' Change "Sheet1" to your sheet name

    ' Hide columns BD to the last possible column
    ws.Rows("161:" & Rows.Count).Hidden = True
    
End Sub

Private Sub Salaire_Change()
Dim ws As Worksheet
    Dim targetCell As Range
    
    ' Set the worksheet and textbox
    Set ws = ThisWorkbook.Sheets("DashBoard") ' Replace with your sheet name
    
    ' Set the target cell where the textbox value should be copied
    Set targetCell = ThisWorkbook.Sheets("Sheet1").Range("B26") ' Replace with your target cell
    
    ' Copy the value from the textbox to the target cell
    targetCell.value = ws.OLEObjects("Salaire").Object.Text
End Sub
