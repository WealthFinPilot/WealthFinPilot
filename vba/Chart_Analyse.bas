Attribute VB_Name = "Chart_Analyse"
Public BDG, SHname, Cbk As String
Public n As Integer
Private Sub AdjustColorCat2()
Dim Chrt As ChartObject
Dim FRow As Range
Dim YRow As Long
Dim i As Integer, y As Integer, A As Integer
Dim SeriesObj As Series
Dim SeriesObjXY As Series
ThisWorkbook.Sheets("DashBoard").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
A = 0
Set Chrt = ThisWorkbook.Sheets("DashBoard").ChartObjects("Chart 7")

On Error Resume Next
Set SeriesObj = Chrt.Chart.SeriesCollection("DataCat2")
On Error GoTo 0
If SeriesObj Is Nothing Then
Exit Sub
End If
On Error Resume Next
Set SeriesObjXY = Chrt.Chart.SeriesCollection("XY")
On Error GoTo 0
For i = 1 To 7
    
        If Not IsError(ThisWorkbook.Sheets("ControlChart").Cells(i, "BK").value) Then
            Set FRow = ThisWorkbook.Sheets("ControlChart").Range("BC1:BC7").Find(What:=ThisWorkbook.Sheets("ControlChart").Cells(i, "BJ").value, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
            
            If Not FRow Is Nothing Then
                YRow = FRow.Row
                ' Ensure YRow is valid
                If YRow >= 1 Then
                        SeriesObj.Points(i).Format.Fill.ForeColor.RGB = ThisWorkbook.Sheets("ControlChart").Cells(YRow, "BR").Interior.color
                        SeriesObj.Points(i).Format.Fill.Solid
 
                   If Not SeriesObjXY.Points(i) Is Nothing Then
                        SeriesObjXY.Points(i).MarkerBackgroundColor = ThisWorkbook.Sheets("ControlChart").Cells(YRow, "BR").Interior.color
                   End If
              End If
         Else
           If ThisWorkbook.Sheets("ControlChart").Cells(i, "BK").value = ThisWorkbook.Sheets("Library").Range("A227").value Then
           A = 1
           SeriesObj.Points(i).Format.Fill.ForeColor.RGB = ThisWorkbook.Sheets("ControlChart").Cells(14, "BJ").Interior.color
           SeriesObj.Points(i).Format.Fill.Solid
           End If
         End If
        End If
Next i
If A = 1 Then
SeriesObjXY.Points(14).MarkerBackgroundColor = ThisWorkbook.Sheets("ControlChart").Cells(14, "BJ").Interior.color
End If
ThisWorkbook.Sheets("DashBoard").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
End Sub
Sub Chart_Dépenses_Income_1()
Dim pt As PivotTable
Set pt = ThisWorkbook.Sheets("Analyse").PivotTables("Dépense_Income")
If pt.CubeFields("[Measures].[%Dépenses/Income]").Position <> 1 Then
pt.CubeFields("[Measures].[%Dépenses/Income]").Position = pt.CubeFields("[Measures].[%Dépenses/Income]").Position - 1
End If
End Sub
Sub Chart_Dépenses_Income_2()
Dim pt As PivotTable
Set pt = ThisWorkbook.Sheets("Analyse").PivotTables("Dépense_Income")
If pt.CubeFields("[Measures].[$Dépenses/Income]").Position <> 1 Then
pt.CubeFields("[Measures].[$Dépenses/Income]").Position = pt.CubeFields("[Measures].[$Dépenses/Income]").Position - 1
End If
End Sub
Sub Chart_Analyse_Dépenses_1()
Dim pt As PivotTable
Set pt = ThisWorkbook.Sheets("Analyse").PivotTables("Analyse_Dépenses")
If pt.CubeFields("[Measures].[Analyse/Total dépense]").Position <> 1 Then
pt.CubeFields("[Measures].[Analyse/Total dépense]").Position = pt.CubeFields("[Measures].[Analyse/Total dépense]").Position - 1
End If
End Sub
Sub Chart_Analyse_Dépenses_2()
Dim pt As PivotTable
Set pt = ThisWorkbook.Sheets("Analyse").PivotTables("Analyse_Dépenses")
If pt.CubeFields("[Measures].[Analyse/Catégorie]").Position <> 1 Then
pt.CubeFields("[Measures].[Analyse/Catégorie]").Position = pt.CubeFields("[Measures].[Analyse/Catégorie]").Position - 1
End If
End Sub

Sub ETDAappli()
Dim pt As PivotTable
Set pt = ThisWorkbook.Sheets("Extrap").PivotTables("DépenseT")
If pt.CubeFields("[Measures].[Analyse/Total Dépenses]").Position <> 1 Then
pt.CubeFields("[Measures].[Analyse/Total Dépenses]").Position = pt.CubeFields("[Measures].[Analyse/Total Dépenses]").Position - 1
ThisWorkbook.Sheets("Extrapolation").ETDA.BackColor = RGB(0, 255, 0)
ThisWorkbook.Sheets("Extrapolation").ETDC.BackColor = -2147483633
End If
End Sub

Sub ETDAmodAppli()
Dim pt As PivotTable
Set pt = ThisWorkbook.Sheets("Extrap").PivotTables("DépenseTmod")
If pt.CubeFields("[Measures].[Analyse modifiée*/Total Dépenses]").Position <> 1 Then
pt.CubeFields("[Measures].[Analyse modifiée*/Total Dépenses]").Position = pt.CubeFields("[Measures].[Analyse modifiée*/Total Dépenses]").Position - 1
ThisWorkbook.Sheets("Extrapolation").ETDAmod.BackColor = RGB(0, 255, 0)
ThisWorkbook.Sheets("Extrapolation").ETDCmod.BackColor = -2147483633
End If
End Sub

Sub ETDCappli()
Dim pt As PivotTable
Set pt = ThisWorkbook.Sheets("Extrap").PivotTables("DépenseT")
If pt.CubeFields("[Measures].[Analyse/Catégories]").Position <> 1 Then
pt.CubeFields("[Measures].[Analyse/Catégories]").Position = pt.CubeFields("[Measures].[Analyse/Catégories]").Position - 1
ThisWorkbook.Sheets("Extrapolation").ETDC.BackColor = RGB(0, 255, 0)
ThisWorkbook.Sheets("Extrapolation").ETDA.BackColor = -2147483633
End If
End Sub

Sub ETDCmodAppli()
Dim pt As PivotTable
Set pt = ThisWorkbook.Sheets("Extrap").PivotTables("DépenseTmod")
If pt.CubeFields("[Measures].[Analyse modifiée*/Catégorie]").Position <> 1 Then
pt.CubeFields("[Measures].[Analyse modifiée*/Catégorie]").Position = pt.CubeFields("[Measures].[Analyse modifiée*/Catégorie]").Position - 1
ThisWorkbook.Sheets("Extrapolation").ETDCmod.BackColor = RGB(0, 255, 0)
ThisWorkbook.Sheets("Extrapolation").ETDAmod.BackColor = -2147483633
End If
End Sub

Sub EDSDAappli()
Dim pt As PivotTable
Set pt = ThisWorkbook.Sheets("Extrap").PivotTables("TEDS")
If pt.CubeFields("[Measures].[$Dépenses/Salaire]").Position <> 1 Then
pt.CubeFields("[Measures].[$Dépenses/Salaire]").Position = pt.CubeFields("[Measures].[$Dépenses/Salaire]").Position - 1
ThisWorkbook.Sheets("Extrapolation").EDSD.BackColor = RGB(0, 255, 0)
ThisWorkbook.Sheets("Extrapolation").EDSP.BackColor = -2147483633
End If
End Sub

Sub EDSPAappli()
Dim pt As PivotTable
Set pt = ThisWorkbook.Sheets("Extrap").PivotTables("TEDS")
If pt.CubeFields("[Measures].[%Dépenses/Salaire]").Position <> 1 Then
pt.CubeFields("[Measures].[%Dépenses/Salaire]").Position = pt.CubeFields("[Measures].[%Dépenses/Salaire]").Position - 1
ThisWorkbook.Sheets("Extrapolation").EDSP.BackColor = RGB(0, 255, 0)
ThisWorkbook.Sheets("Extrapolation").EDSD.BackColor = -2147483633
End If
End Sub
Sub EDSDmodAappli()
Dim pt As PivotTable
Set pt = ThisWorkbook.Sheets("Extrap").PivotTables("TEDSmod")
If pt.CubeFields("[Measures].[$Dépense/Salaire(*Modifié)]").Position <> 1 Then
pt.CubeFields("[Measures].[$Dépense/Salaire(*Modifié)]").Position = pt.CubeFields("[Measures].[$Dépense/Salaire(*Modifié)]").Position - 1
ThisWorkbook.Sheets("Extrapolation").EDSDmod.BackColor = RGB(0, 255, 0)
ThisWorkbook.Sheets("Extrapolation").EDSPmod.BackColor = -2147483633
End If
End Sub

Sub EDSPmodAappli()
Dim pt As PivotTable
Set pt = ThisWorkbook.Sheets("Extrap").PivotTables("TEDSmod")
If pt.CubeFields("[Measures].[%Dépense/Salaire(*Modifié)]").Position <> 1 Then
pt.CubeFields("[Measures].[%Dépense/Salaire(*Modifié)]").Position = pt.CubeFields("[Measures].[%Dépense/Salaire(*Modifié)]").Position - 1
ThisWorkbook.Sheets("Extrapolation").EDSPmod.BackColor = RGB(0, 255, 0)
ThisWorkbook.Sheets("Extrapolation").EDSDmod.BackColor = -2147483633
End If
End Sub

Sub RefreshPivot()
ThisWorkbook.Sheets("Extrap").PivotTables("PivotTable6").RefreshTable
ThisWorkbook.Sheets("Extrap").PivotTables("PivotTable7").RefreshTable
End Sub
Sub ShapeStrat1_Click()
Dim chartObj As ChartObject
Dim series1, series2, series3 As Series
Dim SHP1, SHP2, SHP3, SHP4, SHP5 As Shape
Dim TbWEP As ListObject
Dim dataLabel As dataLabel
Application.ScreenUpdating = False
ThisWorkbook.Sheets("DashBoard").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
Set chartObj = ThisWorkbook.Sheets("DashBoard").ChartObjects("Chart Strategique")
Set series1 = chartObj.Chart.SeriesCollection("Stratégie 1")
Set series2 = chartObj.Chart.SeriesCollection("Stratégie 2")
Set series3 = chartObj.Chart.SeriesCollection("Stratégie 3")
Set dataLabel = series1.Points(2).dataLabel
Set TbWEP = ThisWorkbook.Sheets("WB_Data").ListObjects("TbDataEP")

If TbWEP.ListColumns("Curve").DataBodyRange(1, 1).value = 0 Then
Exit Sub
End If



If TbWEP.ListColumns("CHK").DataBodyRange(1, 1).value = 1 Then
series1.Format.Line.Visible = msoFalse
dataLabel.Format.TextFrame2.TextRange.Font.Fill.Visible = msoFalse ' Removes fill (makes the text transparent)
 
TbWEP.ListColumns("CHK").DataBodyRange(1, 1).value = 0

ElseIf TbWEP.ListColumns("CHK").DataBodyRange(1, 1).value = 0 Then
dataLabel.Format.TextFrame2.TextRange.Font.Fill.Visible = msoTrue
dataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(197, 78, 248)
series1.Format.Line.Visible = msoTrue
series1.Format.Fill.ForeColor.RGB = RGB(197, 78, 248)
series1.MarkerStyle = xlMarkerStyleNone
series2.Format.Line.Visible = msoFalse
series3.Format.Line.Visible = msoFalse
Set dataLabel = series2.Points(2).dataLabel
dataLabel.Format.TextFrame2.TextRange.Font.Fill.Visible = msoFalse
Set dataLabel = series3.Points(2).dataLabel
dataLabel.Format.TextFrame2.TextRange.Font.Fill.Visible = msoFalse
Set SHP1 = ThisWorkbook.Sheets("DashBoard").Shapes("Shape SH1")
Set SHP2 = ThisWorkbook.Sheets("DashBoard").Shapes("Shape SH2")
Set SHP3 = ThisWorkbook.Sheets("DashBoard").Shapes("Shape SH3")
Set SHP4 = ThisWorkbook.Sheets("DashBoard").Shapes("Shape SH4")
Set SHP5 = ThisWorkbook.Sheets("DashBoard").Shapes("Shape SH5")
SHP1.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(197, 78, 248)
SHP2.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(197, 78, 248)
SHP3.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(197, 78, 248)
SHP4.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(197, 78, 248)
SHP5.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(197, 78, 248)
TbWEP.ListColumns("CHK").DataBodyRange(1, 1).value = 1
TbWEP.ListColumns("CHK").DataBodyRange(2, 1).value = 0
TbWEP.ListColumns("CHK").DataBodyRange(3, 1).value = 0
End If
ThisWorkbook.Sheets("DashBoard").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
Application.ScreenUpdating = True
End Sub
Sub ShapeStrat2_Click()
Dim chartObj As ChartObject
Dim series1, series2, series3 As Series
Dim SHP1, SHP2, SHP3, SHP4, SHP5 As Shape
Dim TbWEP As ListObject
Dim dataLabel As dataLabel
Application.ScreenUpdating = False
ThisWorkbook.Sheets("DashBoard").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
Set chartObj = ThisWorkbook.Sheets("DashBoard").ChartObjects("Chart Strategique")
Set series1 = chartObj.Chart.SeriesCollection("Stratégie 1")
Set series2 = chartObj.Chart.SeriesCollection("Stratégie 2")
Set series3 = chartObj.Chart.SeriesCollection("Stratégie 3")
Set dataLabel = series2.Points(2).dataLabel
Set TbWEP = ThisWorkbook.Sheets("WB_Data").ListObjects("TbDataEP")

If TbWEP.ListColumns("Curve").DataBodyRange(2, 1).value = 0 Then
Exit Sub
End If

If TbWEP.ListColumns("CHK").DataBodyRange(2, 1).value = 1 Then
series2.Format.Line.Visible = msoFalse
dataLabel.Format.TextFrame2.TextRange.Font.Fill.Visible = msoFalse
TbWEP.ListColumns("CHK").DataBodyRange(2, 1).value = 0
ElseIf TbWEP.ListColumns("CHK").DataBodyRange(2, 1).value = 0 Then
dataLabel.Format.TextFrame2.TextRange.Font.Fill.Visible = msoTrue
dataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 0)
series2.Format.Line.Visible = msoTrue
series2.Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
series2.MarkerStyle = xlMarkerStyleNone
series1.Format.Line.Visible = msoFalse
series3.Format.Line.Visible = msoFalse
Set dataLabel = series1.Points(2).dataLabel
dataLabel.Format.TextFrame2.TextRange.Font.Fill.Visible = msoFalse
Set dataLabel = series3.Points(2).dataLabel
dataLabel.Format.TextFrame2.TextRange.Font.Fill.Visible = msoFalse
Set SHP1 = ThisWorkbook.Sheets("DashBoard").Shapes("Shape SH1")
Set SHP2 = ThisWorkbook.Sheets("DashBoard").Shapes("Shape SH2")
Set SHP3 = ThisWorkbook.Sheets("DashBoard").Shapes("Shape SH3")
Set SHP4 = ThisWorkbook.Sheets("DashBoard").Shapes("Shape SH4")
Set SHP5 = ThisWorkbook.Sheets("DashBoard").Shapes("Shape SH5")
SHP1.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 0)
SHP2.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 0)
SHP3.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 0)
SHP4.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 0)
SHP5.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 0)

TbWEP.ListColumns("CHK").DataBodyRange(1, 1).value = 0
TbWEP.ListColumns("CHK").DataBodyRange(2, 1).value = 1
TbWEP.ListColumns("CHK").DataBodyRange(3, 1).value = 0
End If
ThisWorkbook.Sheets("DashBoard").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
Application.ScreenUpdating = True
End Sub
Sub ShapeStrat3_Click()
Dim chartObj As ChartObject
Dim series1, series2, series3 As Series
Dim SHP1, SHP2, SHP3, SHP4, SHP5 As Shape
Dim TbWEP As ListObject
Dim dataLabel As dataLabel
Application.ScreenUpdating = False
ThisWorkbook.Sheets("DashBoard").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
Set chartObj = ThisWorkbook.Sheets("DashBoard").ChartObjects("Chart Strategique")
Set series1 = chartObj.Chart.SeriesCollection("Stratégie 1")
Set series2 = chartObj.Chart.SeriesCollection("Stratégie 2")
Set series3 = chartObj.Chart.SeriesCollection("Stratégie 3")
Set dataLabel = series3.Points(2).dataLabel
Set TbWEP = ThisWorkbook.Sheets("WB_Data").ListObjects("TbDataEP")

If TbWEP.ListColumns("Curve").DataBodyRange(3, 1).value = 0 Then
Exit Sub
End If


If TbWEP.ListColumns("CHK").DataBodyRange(3, 1).value = 1 Then
series3.Format.Line.Visible = msoFalse
dataLabel.Format.TextFrame2.TextRange.Font.Fill.Visible = msoFalse
TbWEP.ListColumns("CHK").DataBodyRange(3, 1).value = 0
ElseIf TbWEP.ListColumns("CHK").DataBodyRange(3, 1).value = 0 Then
dataLabel.Format.TextFrame2.TextRange.Font.Fill.Visible = msoTrue
dataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(250, 97, 72)
series3.Format.Line.Visible = msoTrue
series3.Format.Fill.ForeColor.RGB = RGB(250, 97, 72)
series3.MarkerStyle = xlMarkerStyleNone
series1.Format.Line.Visible = msoFalse
series2.Format.Line.Visible = msoFalse
Set dataLabel = series2.Points(2).dataLabel
dataLabel.Format.TextFrame2.TextRange.Font.Fill.Visible = msoFalse
Set dataLabel = series1.Points(2).dataLabel
dataLabel.Format.TextFrame2.TextRange.Font.Fill.Visible = msoFalse
Set SHP1 = ThisWorkbook.Sheets("DashBoard").Shapes("Shape SH1")
Set SHP2 = ThisWorkbook.Sheets("DashBoard").Shapes("Shape SH2")
Set SHP3 = ThisWorkbook.Sheets("DashBoard").Shapes("Shape SH3")
Set SHP4 = ThisWorkbook.Sheets("DashBoard").Shapes("Shape SH4")
Set SHP5 = ThisWorkbook.Sheets("DashBoard").Shapes("Shape SH5")
SHP1.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(250, 97, 72)
SHP2.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(250, 97, 72)
SHP3.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(250, 97, 72)
SHP4.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(250, 97, 72)
SHP5.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(250, 97, 72)

TbWEP.ListColumns("CHK").DataBodyRange(1, 1).value = 0
TbWEP.ListColumns("CHK").DataBodyRange(2, 1).value = 0
TbWEP.ListColumns("CHK").DataBodyRange(3, 1).value = 1
End If
ThisWorkbook.Sheets("DashBoard").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
Application.ScreenUpdating = True
End Sub

Sub ShapeStrat11_Click()
Dim chartObj As ChartObject
Dim series1, series2, series3, series11, series22, series33 As Series
Dim TbWEP As ListObject
Dim dataLabel As dataLabel
Dim SHStrat As Shape
Dim yAxis As Axis
Dim FIRE As Double
Application.ScreenUpdating = False
ThisWorkbook.Sheets("DashBoard").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value

Set chartObj = ThisWorkbook.Sheets("DashBoard").ChartObjects("Chart Strategique")
Set series1 = chartObj.Chart.SeriesCollection("Stratégie 1")
Set series11 = chartObj.Chart.SeriesCollection("Stratégie1")
Set dataLabel = series1.Points(2).dataLabel
Set TbWEP = ThisWorkbook.Sheets("WB_Data").ListObjects("TbDataEP")
Set SHStrat = ThisWorkbook.Sheets("DashBoard").Shapes("ShapeStrat11")


If TbWEP.ListColumns("Curve").DataBodyRange(1, 1).value = 1 Then
series1.Format.Line.Visible = msoFalse
series11.Format.Line.Visible = msoFalse
SHStrat.Line.ForeColor.RGB = RGB(118, 113, 113) ' Red border
SHStrat.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(118, 113, 113)
dataLabel.Format.TextFrame2.TextRange.Font.Fill.Visible = msoFalse ' Removes fill (makes the text transparent)
TbWEP.ListColumns("CHK").DataBodyRange(1, 1).value = 0
TbWEP.ListColumns("Curve").DataBodyRange(1, 1).value = 0
ElseIf TbWEP.ListColumns("Curve").DataBodyRange(1, 1).value = 0 Then
SHStrat.Line.ForeColor.RGB = RGB(170, 220, 215) ' Red border
SHStrat.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(170, 220, 215)
dataLabel.Format.TextFrame2.TextRange.Font.Fill.Visible = msoTrue
dataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(197, 78, 248)
series11.Format.Line.Visible = msoTrue
series11.MarkerStyle = xlMarkerStyleNone ' Remove markers
series11.Format.Line.Weight = 1 ' Set line width to 1.25
series11.Smooth = True ' Enable smoothed line
series11.Format.Fill.ForeColor.RGB = RGB(197, 78, 248)
series1.Format.Line.Visible = msoTrue
series1.Format.Fill.ForeColor.RGB = RGB(197, 78, 248)
series1.MarkerStyle = xlMarkerStyleNone
TbWEP.ListColumns("Curve").DataBodyRange(1, 1).value = 1

If TbWEP.ListColumns("Curve").DataBodyRange(2, 1).value = 0 And TbWEP.ListColumns("Curve").DataBodyRange(3, 1).value = 0 Then
FIRE = TbWEP.ListColumns("FIRE").DataBodyRange(1, 1).value * 1.2
Set yAxis = chartObj.Chart.Axes(xlValue)
yAxis.MaximumScale = FIRE
End If



TbWEP.ListColumns("CHK").DataBodyRange(1, 1).value = 1
TbWEP.ListColumns("CHK").DataBodyRange(2, 1).value = 0
TbWEP.ListColumns("CHK").DataBodyRange(3, 1).value = 0
End If
ThisWorkbook.Sheets("DashBoard").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
Application.ScreenUpdating = True
End Sub

Sub ShapeStrat22_Click()
Dim chartObj As ChartObject
Dim series1, series2, series3, series11, series22, series33 As Series
Dim TbWEP As ListObject
Dim dataLabel As dataLabel
Dim SHStrat As Shape
Dim yAxis As Axis
Dim FIRE As Double
Application.ScreenUpdating = False
ThisWorkbook.Sheets("DashBoard").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
Set chartObj = ThisWorkbook.Sheets("DashBoard").ChartObjects("Chart Strategique")
Set series2 = chartObj.Chart.SeriesCollection("Stratégie 2")
Set series22 = chartObj.Chart.SeriesCollection("Stratégie2")
Set dataLabel = series2.Points(2).dataLabel
Set TbWEP = ThisWorkbook.Sheets("WB_Data").ListObjects("TbDataEP")
Set SHStrat = ThisWorkbook.Sheets("DashBoard").Shapes("ShapeStrat22")

If TbWEP.ListColumns("Curve").DataBodyRange(2, 1).value = 1 Then
series2.Format.Line.Visible = msoFalse
series22.Format.Line.Visible = msoFalse
SHStrat.Line.ForeColor.RGB = RGB(118, 113, 113) ' Red border
SHStrat.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(118, 113, 113)
dataLabel.Format.TextFrame2.TextRange.Font.Fill.Visible = msoFalse ' Removes fill (makes the text transparent)
TbWEP.ListColumns("CHK").DataBodyRange(2, 1).value = 0
TbWEP.ListColumns("Curve").DataBodyRange(2, 1).value = 0


If TbWEP.ListColumns("Curve").DataBodyRange(1, 1).value = 1 And TbWEP.ListColumns("Curve").DataBodyRange(3, 1).value = 0 Then
FIRE = TbWEP.ListColumns("FIRE").DataBodyRange(1, 1).value * 1.2
Set yAxis = chartObj.Chart.Axes(xlValue)
yAxis.MaximumScale = FIRE
End If


ElseIf TbWEP.ListColumns("Curve").DataBodyRange(2, 1).value = 0 Then
SHStrat.Line.ForeColor.RGB = RGB(170, 220, 215) ' Red border
SHStrat.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(170, 220, 215)
dataLabel.Format.TextFrame2.TextRange.Font.Fill.Visible = msoTrue
dataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 0)
series22.Format.Line.Visible = msoTrue
series22.MarkerStyle = xlMarkerStyleNone ' Remove markers
series22.Format.Line.Weight = 1 ' Set line width to 1.25
series22.Smooth = True ' Enable smoothed line
series22.Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
series2.Format.Line.Visible = msoTrue
series2.Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
series2.MarkerStyle = xlMarkerStyleNone
TbWEP.ListColumns("Curve").DataBodyRange(2, 1).value = 1

If TbWEP.ListColumns("Curve").DataBodyRange(3, 1).value = 0 Then
FIRE = TbWEP.ListColumns("FIRE").DataBodyRange(2, 1).value * 1.2
Set yAxis = chartObj.Chart.Axes(xlValue)
yAxis.MaximumScale = FIRE
End If

TbWEP.ListColumns("CHK").DataBodyRange(1, 1).value = 0
TbWEP.ListColumns("CHK").DataBodyRange(2, 1).value = 1
TbWEP.ListColumns("CHK").DataBodyRange(3, 1).value = 0
End If
ThisWorkbook.Sheets("DashBoard").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
Application.ScreenUpdating = True
End Sub

Sub ShapeStrat33_Click()
Dim chartObj As ChartObject
Dim series1, series2, series3, series11, series22, series33 As Series
Dim TbWEP As ListObject
Dim dataLabel As dataLabel
Dim SHStrat As Shape
Dim yAxis As Axis
Dim FIRE As Double
Application.ScreenUpdating = False
ThisWorkbook.Sheets("DashBoard").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
Set chartObj = ThisWorkbook.Sheets("DashBoard").ChartObjects("Chart Strategique")
Set series3 = chartObj.Chart.SeriesCollection("Stratégie 3")
Set series33 = chartObj.Chart.SeriesCollection("Stratégie3")
Set dataLabel = series3.Points(2).dataLabel
Set TbWEP = ThisWorkbook.Sheets("WB_Data").ListObjects("TbDataEP")
Set SHStrat = ThisWorkbook.Sheets("DashBoard").Shapes("ShapeStrat33")


If TbWEP.ListColumns("Curve").DataBodyRange(3, 1).value = 1 Then
series3.Format.Line.Visible = msoFalse
series33.Format.Line.Visible = msoFalse
SHStrat.Line.ForeColor.RGB = RGB(118, 113, 113) ' Red border
SHStrat.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(118, 113, 113)
dataLabel.Format.TextFrame2.TextRange.Font.Fill.Visible = msoFalse ' Removes fill (makes the text transparent)
TbWEP.ListColumns("CHK").DataBodyRange(3, 1).value = 0
TbWEP.ListColumns("Curve").DataBodyRange(3, 1).value = 0

If TbWEP.ListColumns("Curve").DataBodyRange(2, 1).value = 1 Then
FIRE = TbWEP.ListColumns("FIRE").DataBodyRange(2, 1).value * 1.2
Set yAxis = chartObj.Chart.Axes(xlValue)
yAxis.MaximumScale = FIRE
End If
If TbWEP.ListColumns("Curve").DataBodyRange(1, 1).value = 1 And TbWEP.ListColumns("Curve").DataBodyRange(2, 1).value = 0 Then
FIRE = TbWEP.ListColumns("FIRE").DataBodyRange(1, 1).value * 1.2
Set yAxis = chartObj.Chart.Axes(xlValue)
yAxis.MaximumScale = FIRE
End If


ElseIf TbWEP.ListColumns("Curve").DataBodyRange(3, 1).value = 0 Then
SHStrat.Line.ForeColor.RGB = RGB(170, 220, 215) ' Red border
SHStrat.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(170, 220, 215)
dataLabel.Format.TextFrame2.TextRange.Font.Fill.Visible = msoTrue
dataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(250, 97, 72)
series33.Format.Line.Visible = msoTrue
series33.MarkerStyle = xlMarkerStyleNone ' Remove markers
series33.Format.Line.Weight = 1 ' Set line width to 1.25
series33.Smooth = True ' Enable smoothed line
series33.Format.Fill.ForeColor.RGB = RGB(250, 97, 72)

    series3.Format.Line.Visible = msoTrue
    series3.Format.Fill.ForeColor.RGB = RGB(250, 97, 72)
    series3.MarkerStyle = xlMarkerStyleNone

FIRE = TbWEP.ListColumns("FIRE").DataBodyRange(3, 1).value
FIRE = FIRE * 1.2
Set yAxis = chartObj.Chart.Axes(xlValue)
yAxis.MaximumScale = FIRE

TbWEP.ListColumns("Curve").DataBodyRange(3, 1).value = 1

TbWEP.ListColumns("CHK").DataBodyRange(1, 1).value = 0
TbWEP.ListColumns("CHK").DataBodyRange(2, 1).value = 0
TbWEP.ListColumns("CHK").DataBodyRange(3, 1).value = 1
End If
ThisWorkbook.Sheets("DashBoard").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
Application.ScreenUpdating = True
End Sub

Sub SETTING_EP()
Application.ScreenUpdating = False
ThisWorkbook.Sheets("ControlChart").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
SETEP.Show
ThisWorkbook.Sheets("ControlChart").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("DashBoard").Activate
Application.ScreenUpdating = True
End Sub
Sub SETTING_CELI()
CELIF.Show
End Sub
Sub SETTING_NetIncome()
NetIncome.Show
End Sub
Sub SETTING_SETDATA()
Application.ScreenUpdating = False
ThisWorkbook.Sheets("WB_Data").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Budget1").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Budget2").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("TamponBudget1").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("TamponBudget2").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ChangeCAT.Show
ThisWorkbook.Sheets("WB_Data").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Budget1").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Budget2").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("TamponBudget1").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("TamponBudget2").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("DashBoard").Activate
Application.ScreenUpdating = True
End Sub
Sub ENTER_DATA()
Dim tbAn As ListObject
Dim result As VbMsgBoxResult
Dim An As Integer
Dim ANRow As Long
Dim Mois, LibR As String
Dim FRow As Range
Application.ScreenUpdating = False
ThisWorkbook.Sheets("WB_Data").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value


Set tbAn = ThisWorkbook.Sheets("WB_Data").ListObjects("Year")
An = VBA.DateTime.Year(Date)

                Set FRow = tbAn.ListColumns("Année").DataBodyRange.Find(What:=An, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                 On Error Resume Next
                 If Not FRow Is Nothing Then
                 Else
                 LibR = ThisWorkbook.Sheets("Library").Range("A9").value
                 result = MsgBox(LibR, vbYesNo)
                 Select Case result
                        Case Is = vbYes
                        ANRow = tbAn.ListRows.Count
                        tbAn.ListColumns("Année").DataBodyRange(ANRow + 1, 1).value = An
                        
                        Case Is = vbNo
                        ThisWorkbook.Sheets("WB_Data").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
                        ThisWorkbook.Sheets("Input").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
                        ThisWorkbook.Sheets("Data Base").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
                        ThisWorkbook.Sheets("DashBoard").Activate
                        Application.ScreenUpdating = True
                        Exit Sub
                End Select
                End If


InputData.Show


ThisWorkbook.Sheets("WB_Data").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("DashBoard").Activate
Application.ScreenUpdating = True
End Sub
Sub AUTOMATE_DATA()
Dim tbAn As ListObject
Dim result As VbMsgBoxResult
Dim An As Integer
Dim ANRow As Long
Dim Mois, LibR As String
Dim FRow As Range
Dim tbCCY As ListObject
Application.ScreenUpdating = False
ThisWorkbook.Sheets("WB_Data").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value


Set tbAn = ThisWorkbook.Sheets("WB_Data").ListObjects("Year")
An = VBA.DateTime.Year(Date)

                Set FRow = tbAn.ListColumns("Année").DataBodyRange.Find(What:=An, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
                 On Error Resume Next
                 If Not FRow Is Nothing Then
                 Else
                 LibR = ThisWorkbook.Sheets("Library").Range("A9").value
                 result = MsgBox(LibR, vbYesNo)
                 Select Case result
                        Case Is = vbYes
                        ANRow = tbAn.ListRows.Count
                        tbAn.ListColumns("Année").DataBodyRange(ANRow + 1, 1).value = An
                        
                        Case Is = vbNo
                        ThisWorkbook.Sheets("Data Entry").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
                        ThisWorkbook.Sheets("WB_Data").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
                        ThisWorkbook.Sheets("Input").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
                        ThisWorkbook.Sheets("Data Base").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
                        ThisWorkbook.Sheets("DashBoard").Activate
                        Application.ScreenUpdating = True
                        Exit Sub
                End Select
                End If

InputData.Show
Set tbCCY = ThisWorkbook.Sheets("WB_Data").ListObjects("tbCCY")
tbCCY.ListColumns("Path").DataBodyRange(4, 1).value = ""
ThisWorkbook.Sheets("Data Entry").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("WB_Data").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("DashBoard").Activate
Application.ScreenUpdating = True
End Sub


Sub Analyse_Data()
Dim Sarray, emptyCells As Range

Application.ScreenUpdating = False
ThisWorkbook.Sheets("WB_Data").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Budget1").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Budget2").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value


Set tb1 = ThisWorkbook.Sheets("Input").ListObjects("DataTable")
Set Sarray = tb1.ListColumns("Check").DataBodyRange
On Error Resume Next
Set emptyCells = Sarray.SpecialCells(xlCellTypeBlanks)
On Error GoTo 0

   If Not emptyCells Is Nothing Then
        Set FindFirstEmptyCellInColumn = emptyCells.Cells(1)
            Start = FindFirstEmptyCellInColumn.Row - 1
            Tri_Data
            Chart_Analyse.AdjustColorCat2
   End If
   
On Error Resume Next

ThisWorkbook.RefreshAll

ThisWorkbook.Sheets("WB_Data").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Budget1").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Budget2").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("DashBoard").Activate
Application.ScreenUpdating = True
End Sub
Sub Budgetisation()
Dim result As VbMsgBoxResult
Dim LibR As String
Application.ScreenUpdating = False
ThisWorkbook.Sheets("WB_Data").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
LibR = ThisWorkbook.Sheets("Library").Range("A8").value
result = MsgBox(LibR, vbYesNo)
Select Case result
       Case Is = vbYes
       tampon
       Case Is = vbNo
End Select

On Error Resume Next


ThisWorkbook.Sheets("WB_Data").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("DashBoard").Activate
Application.ScreenUpdating = True
End Sub

Sub ActiveBudget1()
Dim shp, shpOV, shpBO1 As Shape
Dim rng As Range
Dim TbBudget As ListObject
Application.ScreenUpdating = False
ThisWorkbook.Sheets("DashBoard").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value

Set shp = ThisWorkbook.Sheets("DashBoard").Shapes("Group 5829")
Set shpOV = ThisWorkbook.Sheets("DashBoard").Shapes("Group 5437")
Set shpBO1 = ThisWorkbook.Sheets("DashBoard").Shapes("OVALB1")
Set TbBudget = ThisWorkbook.Sheets("WB_Data").ListObjects("Budget")

If TbBudget.ListColumns("Activate").DataBodyRange(1, 1).value = 0 Then
   TbBudget.ListColumns("Activate").DataBodyRange(1, 1).value = 1
   shp.Visible = msoTrue
   shpOV.Visible = msoTrue
   shpBO1.Visible = msoTrue
ElseIf TbBudget.ListColumns("Activate").DataBodyRange(1, 1).value = 1 Then
   TbBudget.ListColumns("Activate").DataBodyRange(1, 1).value = 0
   shp.Visible = msoFalse
   shpOV.Visible = msoFalse
   shpBO1.Visible = msoFalse
End If
    
ThisWorkbook.Sheets("DashBoard").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
Application.ScreenUpdating = True
End Sub
Sub ActiveBudget2()
Dim shp, shpOV, shpBO2 As Shape
Dim rng As Range
Dim TbBudget As ListObject

Application.ScreenUpdating = False
ThisWorkbook.Sheets("DashBoard").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
Set shp = ThisWorkbook.Sheets("DashBoard").Shapes("Group 5848")
Set shpOV = ThisWorkbook.Sheets("DashBoard").Shapes("Group 5438")
Set shpBO2 = ThisWorkbook.Sheets("DashBoard").Shapes("OVALB2")
Set TbBudget = ThisWorkbook.Sheets("WB_Data").ListObjects("Budget")

If TbBudget.ListColumns("Activate").DataBodyRange(2, 1).value = 0 Then
   TbBudget.ListColumns("Activate").DataBodyRange(2, 1).value = 1
   shp.Visible = msoTrue
   shpOV.Visible = msoTrue
   shpBO2.Visible = msoTrue
ElseIf TbBudget.ListColumns("Activate").DataBodyRange(2, 1).value = 1 Then
   TbBudget.ListColumns("Activate").DataBodyRange(2, 1).value = 0
   shp.Visible = msoFalse
   shpOV.Visible = msoFalse
   shpBO2.Visible = msoFalse
End If
    
ThisWorkbook.Sheets("DashBoard").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
Application.ScreenUpdating = True
End Sub
Sub RESET_DATA()
Dim result As VbMsgBoxResult
Dim newDataBodyRange As Range
Dim LibR As String
LibR = ThisWorkbook.Sheets("Library").Range("A11").value
result = MsgBox(LibR, vbYesNo)
       
       If result = vbYes Then
            Application.ScreenUpdating = True
            ThisWorkbook.Sheets("Input").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
            ThisWorkbook.Sheets("Input").Range("Z1:Z" & ThisWorkbook.Sheets("Input").Rows.Count).ClearContents
            ThisWorkbook.Sheets("Input").Range("Y1:Y" & ThisWorkbook.Sheets("Input").Rows.Count).ClearContents
            Set tb1 = ThisWorkbook.Sheets("Input").ListObjects("DataTable")

            tb1.DataBodyRange.Clear
            tb1.ListColumns("Année").DataBodyRange(0 + 1, 1).value = "Exemple"
            tb1.ListColumns("Année").DataBodyRange(1, 1).Font.color = RGB(255, 255, 255)
            Set newDataBodyRange = tb1.ListRows(1).Range
            tb1.Resize tb1.Range.Resize(newDataBodyRange.Rows.Count + 1, newDataBodyRange.Columns.Count)
            ThisWorkbook.Sheets("Input").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
            ThisWorkbook.Sheets("DashBoard").Activate
            Application.ScreenUpdating = True
        End If
End Sub
Sub Open_EPCIBLE()
Epargne.Show
End Sub
Sub LauchEPProgress()
Application.ScreenUpdating = False
ThisWorkbook.Sheets("ControlChart").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
EPProgress.Show
ThisWorkbook.Sheets("ControlChart").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("DashBoard").Activate
Application.ScreenUpdating = True
End Sub
Sub GoBudget2()
Dim result As VbMsgBoxResult
Dim LibR As String
Application.ScreenUpdating = False
ThisWorkbook.Sheets("WB_Data").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
LibR = ThisWorkbook.Sheets("Library").Range("A8").value
result = MsgBox(LibR, vbYesNo)
Select Case result
       Case Is = vbYes
       tampon
       ThisWorkbook.RefreshAll
       Case Is = vbNo
End Select

On Error Resume Next


ThisWorkbook.Sheets("WB_Data").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("DashBoard").Activate
Application.ScreenUpdating = True
End Sub
Sub GoBudget1()

Dim result As VbMsgBoxResult
Dim LibR As String
Application.ScreenUpdating = False
ThisWorkbook.Sheets("WB_Data").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Unprotect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
LibR = ThisWorkbook.Sheets("Library").Range("A8").value
result = MsgBox(LibR, vbYesNo)
Select Case result
       Case Is = vbYes
       tampon
       ThisWorkbook.RefreshAll
       Case Is = vbNo
End Select

On Error Resume Next


ThisWorkbook.Sheets("WB_Data").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Input").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("Data Base").Protect Password:=ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead").ListColumns("MP").DataBodyRange(1, 1).value
ThisWorkbook.Sheets("DashBoard").Activate
Application.ScreenUpdating = True
End Sub
Sub LaunchBudget2()
Dim TbB As ListObject
Set TbB = ThisWorkbook.Sheets("WB_Data").ListObjects("Budget")
TbB.ListColumns("Input").DataBodyRange(2, 1).value = 1
TbB.ListColumns("Input").DataBodyRange(1, 1).value = 0
BDG = "Budget2"
Tamp = "Tampon" & BDG
Call Moyenner(Tamp, BDG)
End Sub
Sub LaunchBudget1()
Dim TbB As ListObject
Set TbB = ThisWorkbook.Sheets("WB_Data").ListObjects("Budget")
TbB.ListColumns("Input").DataBodyRange(1, 1).value = 1
TbB.ListColumns("Input").DataBodyRange(2, 1).value = 0

BDG = "Budget1"
Tamp = "Tampon" & BDG
Call Moyenner(Tamp, BDG)
End Sub

Sub ShapeI11()
 For n = 1 To 3
    Cbk = "I" & 1 & n
    SHname = "SH " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 1 Then
    ThisWorkbook.Sheets("ControlChart_Budget1").Cells(1 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(250, 97, 72)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(250, 97, 72)
    End If
 Next n
End Sub
Sub ShapeI12()
 For n = 1 To 3
    Cbk = "I" & 1 & n
    SHname = "SH " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 2 Then
    ThisWorkbook.Sheets("ControlChart_Budget1").Cells(1 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(250, 97, 72)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(250, 97, 72)
    End If
 Next n
End Sub
Sub ShapeI13()
 For n = 1 To 3
    Cbk = "I" & 1 & n
    SHname = "SH " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 3 Then
    ThisWorkbook.Sheets("ControlChart_Budget1").Cells(1 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(250, 97, 72)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(250, 97, 72)
    End If
 Next n
End Sub
Sub ShapeI21()

 For n = 1 To 3
    Cbk = "I" & 2 & n
    SHname = "SH " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 1 Then
    ThisWorkbook.Sheets("ControlChart_Budget1").Cells(2 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(250, 97, 72)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(250, 97, 72)
    End If
 Next n
End Sub
Sub ShapeI22()
 For n = 1 To 3
    Cbk = "I" & 2 & n
    SHname = "SH " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 2 Then
    ThisWorkbook.Sheets("ControlChart_Budget1").Cells(2 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(250, 97, 72)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(250, 97, 72)
    End If
 Next n
End Sub
Sub ShapeI23()
 For n = 1 To 3
    Cbk = "I" & 2 & n
    SHname = "SH " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 3 Then
    ThisWorkbook.Sheets("ControlChart_Budget1").Cells(2 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(250, 97, 72)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(250, 97, 72)
    End If
 Next n
End Sub
Sub ShapeI31()
 For n = 1 To 3
    Cbk = "I" & 3 & n
    SHname = "SH " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 1 Then
    ThisWorkbook.Sheets("ControlChart_Budget1").Cells(3 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(250, 97, 72)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(250, 97, 72)
    End If
 Next n
End Sub
Sub ShapeI32()
 For n = 1 To 3
    Cbk = "I" & 3 & n
    SHname = "SH " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 2 Then
    ThisWorkbook.Sheets("ControlChart_Budget1").Cells(3 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(250, 97, 72)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(250, 97, 72)
    End If
 Next n
End Sub
Sub ShapeI33()
 For n = 1 To 3
    Cbk = "I" & 3 & n
    SHname = "SH " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 3 Then
    ThisWorkbook.Sheets("ControlChart_Budget1").Cells(3 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(250, 97, 72)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(250, 97, 72)
    End If
 Next n
End Sub
Sub ShapeI41()
 For n = 1 To 3
    Cbk = "I" & 4 & n
    SHname = "SH " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 1 Then
    ThisWorkbook.Sheets("ControlChart_Budget1").Cells(4 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(250, 97, 72)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(250, 97, 72)
    End If
 Next n
End Sub
Sub ShapeI42()
 For n = 1 To 3
    Cbk = "I" & 4 & n
    SHname = "SH " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 2 Then
    ThisWorkbook.Sheets("ControlChart_Budget1").Cells(4 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(250, 97, 72)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(250, 97, 72)
    End If
 Next n
End Sub
Sub ShapeI43()
 For n = 1 To 3
    Cbk = "I" & 4 & n
    SHname = "SH " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 3 Then
    ThisWorkbook.Sheets("ControlChart_Budget1").Cells(4 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(250, 97, 72)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(250, 97, 72)
    End If
 Next n
End Sub
Sub ShapeI51()
 For n = 1 To 3
    Cbk = "I" & 5 & n
    SHname = "SH " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 1 Then
    ThisWorkbook.Sheets("ControlChart_Budget1").Cells(5 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(250, 97, 72)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(250, 97, 72)
    End If
 Next n
End Sub
Sub ShapeI52()
 For n = 1 To 3
    Cbk = "I" & 5 & n
    SHname = "SH " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 2 Then
    ThisWorkbook.Sheets("ControlChart_Budget1").Cells(5 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(250, 97, 72)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(250, 97, 72)
    End If
 Next n
End Sub
Sub ShapeI53()
 For n = 1 To 3
    Cbk = "I" & 5 & n
    SHname = "SH " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 3 Then
    ThisWorkbook.Sheets("ControlChart_Budget1").Cells(5 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(250, 97, 72)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(250, 97, 72)
    End If
 Next n
End Sub
Sub ShapeI61()
 For n = 1 To 3
    Cbk = "I" & 6 & n
    SHname = "SH " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 1 Then
    ThisWorkbook.Sheets("ControlChart_Budget1").Cells(6 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(250, 97, 72)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(250, 97, 72)
    End If
 Next n
End Sub
Sub ShapeI62()
 For n = 1 To 3
    Cbk = "I" & 6 & n
    SHname = "SH " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 2 Then
    ThisWorkbook.Sheets("ControlChart_Budget1").Cells(6 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(250, 97, 72)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(250, 97, 72)
    End If
 Next n
End Sub
Sub ShapeI63()
 For n = 1 To 3
    Cbk = "I" & 6 & n
    SHname = "SH " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 3 Then
    ThisWorkbook.Sheets("ControlChart_Budget1").Cells(6 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(250, 97, 72)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(250, 97, 72)
    End If
 Next n
End Sub
Sub ShapeI71()
 For n = 1 To 3
    Cbk = "I" & 7 & n
    SHname = "SH " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 1 Then
    ThisWorkbook.Sheets("ControlChart_Budget1").Cells(7 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(250, 97, 72)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(250, 97, 72)
    End If
 Next n
End Sub
Sub ShapeI72()
 For n = 1 To 3
    Cbk = "I" & 7 & n
    SHname = "SH " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 2 Then
    ThisWorkbook.Sheets("ControlChart_Budget1").Cells(7 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(250, 97, 72)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(250, 97, 72)
    End If
 Next n
End Sub
Sub ShapeI73()
 For n = 1 To 3
    Cbk = "I" & 7 & n
    SHname = "SH " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 3 Then
    ThisWorkbook.Sheets("ControlChart_Budget1").Cells(7 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(250, 97, 72)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(250, 97, 72)
    End If
 Next n
End Sub


Sub ShapeAI11()
 For n = 1 To 3
    Cbk = "I" & 1 & n
    SHname = "SHA " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 1 Then
    ThisWorkbook.Sheets("ControlChart_Budget2").Cells(1 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(197, 78, 248)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(197, 78, 248)
    End If
 Next n
End Sub
Sub ShapeAI12()
 For n = 1 To 3
    Cbk = "I" & 1 & n
    SHname = "SHA " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 2 Then
    ThisWorkbook.Sheets("ControlChart_Budget2").Cells(1 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(197, 78, 248)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(197, 78, 248)
    End If
 Next n
End Sub
Sub ShapeAI13()
 For n = 1 To 3
    Cbk = "I" & 1 & n
    SHname = "SHA " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 3 Then
    ThisWorkbook.Sheets("ControlChart_Budget2").Cells(1 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(197, 78, 248)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(197, 78, 248)
    End If
 Next n
End Sub
Sub ShapeAI21()

 For n = 1 To 3
    Cbk = "I" & 2 & n
    SHname = "SHA " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 1 Then
    ThisWorkbook.Sheets("ControlChart_Budget2").Cells(2 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(197, 78, 248)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(197, 78, 248)
    End If
 Next n
End Sub
Sub ShapeAI22()
 For n = 1 To 3
    Cbk = "I" & 2 & n
    SHname = "SHA " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 2 Then
    ThisWorkbook.Sheets("ControlChart_Budget2").Cells(2 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(197, 78, 248)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(197, 78, 248)
    End If
 Next n
End Sub
Sub ShapeAI23()
 For n = 1 To 3
    Cbk = "I" & 2 & n
    SHname = "SHA " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 3 Then
    ThisWorkbook.Sheets("ControlChart_Budget2").Cells(2 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(197, 78, 248)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(197, 78, 248)
    End If
 Next n
End Sub
Sub ShapeAI31()
 For n = 1 To 3
    Cbk = "I" & 3 & n
    SHname = "SHA " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 1 Then
    ThisWorkbook.Sheets("ControlChart_Budget2").Cells(3 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(197, 78, 248)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(197, 78, 248)
    End If
 Next n
End Sub
Sub ShapeAI32()
 For n = 1 To 3
    Cbk = "I" & 3 & n
    SHname = "SHA " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 2 Then
    ThisWorkbook.Sheets("ControlChart_Budget2").Cells(3 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(197, 78, 248)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(197, 78, 248)
    End If
 Next n
End Sub
Sub ShapeAI33()
 For n = 1 To 3
    Cbk = "I" & 3 & n
    SHname = "SHA " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 3 Then
    ThisWorkbook.Sheets("ControlChart_Budget2").Cells(3 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(197, 78, 248)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(197, 78, 248)
    End If
 Next n
End Sub
Sub ShapeAI41()
 For n = 1 To 3
    Cbk = "I" & 4 & n
    SHname = "SHA " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 1 Then
    ThisWorkbook.Sheets("ControlChart_Budget2").Cells(4 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(197, 78, 248)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(197, 78, 248)
    End If
 Next n
End Sub
Sub ShapeAI42()
 For n = 1 To 3
    Cbk = "I" & 4 & n
    SHname = "SHA " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 2 Then
    ThisWorkbook.Sheets("ControlChart_Budget2").Cells(4 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(197, 78, 248)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(197, 78, 248)
    End If
 Next n
End Sub
Sub ShapeAI43()
 For n = 1 To 3
    Cbk = "I" & 4 & n
    SHname = "SHA " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 3 Then
    ThisWorkbook.Sheets("ControlChart_Budget2").Cells(4 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(197, 78, 248)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(197, 78, 248)
    End If
 Next n
End Sub
Sub ShapeAI51()
 For n = 1 To 3
    Cbk = "I" & 5 & n
    SHname = "SHA " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 1 Then
    ThisWorkbook.Sheets("ControlChart_Budget2").Cells(5 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(197, 78, 248)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(197, 78, 248)
    End If
 Next n
End Sub
Sub ShapeAI52()
 For n = 1 To 3
    Cbk = "I" & 5 & n
    SHname = "SHA " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 2 Then
    ThisWorkbook.Sheets("ControlChart_Budget2").Cells(5 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(197, 78, 248)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(197, 78, 248)
    End If
 Next n
End Sub
Sub ShapeAI53()
 For n = 1 To 3
    Cbk = "I" & 5 & n
    SHname = "SHA " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 3 Then
    ThisWorkbook.Sheets("ControlChart_Budget2").Cells(5 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(197, 78, 248)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(197, 78, 248)
    End If
 Next n
End Sub
Sub ShapeAI61()
 For n = 1 To 3
    Cbk = "I" & 6 & n
    SHname = "SHA " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 1 Then
    ThisWorkbook.Sheets("ControlChart_Budget2").Cells(6 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(197, 78, 248)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(197, 78, 248)
    End If
 Next n
End Sub
Sub ShapeAI62()
 For n = 1 To 3
    Cbk = "I" & 6 & n
    SHname = "SHA " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 2 Then
    ThisWorkbook.Sheets("ControlChart_Budget2").Cells(6 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(197, 78, 248)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(197, 78, 248)
    End If
 Next n
End Sub
Sub ShapeAI63()
 For n = 1 To 3
    Cbk = "I" & 6 & n
    SHname = "SHA " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 3 Then
    ThisWorkbook.Sheets("ControlChart_Budget2").Cells(6 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(197, 78, 248)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(197, 78, 248)
    End If
 Next n
End Sub
Sub ShapeAI71()
 For n = 1 To 3
    Cbk = "I" & 7 & n
    SHname = "SHA " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 1 Then
    ThisWorkbook.Sheets("ControlChart_Budget2").Cells(7 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(197, 78, 248)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(197, 78, 248)
    End If
 Next n
End Sub
Sub ShapeAI72()
 For n = 1 To 3
    Cbk = "I" & 7 & n
    SHname = "SHA " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 2 Then
    ThisWorkbook.Sheets("ControlChart_Budget2").Cells(7 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(197, 78, 248)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(197, 78, 248)
    End If
 Next n
End Sub
Sub ShapeAI73()
 For n = 1 To 3
    Cbk = "I" & 7 & n
    SHname = "SHA " & Cbk
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(59, 56, 56)
    If n = 3 Then
    ThisWorkbook.Sheets("ControlChart_Budget2").Cells(7 + 3, 5).value = n
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Fill.ForeColor.RGB = RGB(197, 78, 248)
    ThisWorkbook.Sheets("DashBoard").Shapes(SHname).Line.ForeColor.RGB = RGB(197, 78, 248)
    End If
 Next n
End Sub
