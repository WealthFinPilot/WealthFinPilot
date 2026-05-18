VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SETEP 
   Caption         =   "Plannifiez vos stratégies"
   ClientHeight    =   6540
   ClientLeft      =   112
   ClientTop       =   448
   ClientWidth     =   12272
   OleObjectBlob   =   "SETEP.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "SETEP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TbWEP, DTEP, tbh, tbCCY As ListObject
Public SAB, m, EPM, r1, r2, r3, YrY, FIRE_Number1, FIRE_Number2, FIRE_Number3, Withdrawal, Tolerance, Deposit, MonthlyDeposit, CurrentWithdrawal, Capital1, Capital2, Capital3 As Double
Public A, Q, RT, CT, FT, t, BC, i, j, ER, MaxIterations, IterationCount, n, CheckYear, MissingYears1, MissingYears2, MissingYears3 As Integer
Public VSB, Vm, VEPM, VS1, VS2, VS3, VS4, VRIF, VWM As Variant
Public MySB, Mym, MyEPM, MyS1, MyS2, MyS3, MyS4, MyRIF, MyWM As String
Public FRow As Range
Public Année, CRY, localDec, altDec, vbaSep As String
Public chartObj As ChartObject
Public dataLabel As dataLabel
Public series1, series2, series3, series11, series22, series33 As Series
Public xAxis, yAxis As Axis
Public SHStrat As Shape


Private Sub CheckInf_Click()
If Me.CheckInf.value = True Then
   Me.CheckInf.ForeColor = &HFFFFC0
   Me.LMEP.ForeColor = &HFFFFC0
   Me.LWM.ForeColor = &HFFFFC0
ElseIf Me.CheckInf.value = False Then
   Me.CheckInf.ForeColor = &HC0C0C0
   Me.LMEP.ForeColor = &HC0C0C0
   Me.LWM.ForeColor = &HC0C0C0
End If
End Sub

Private Sub CHKWM_Click()

If Me.CHKWM.value = True Then
   Me.WM.BorderStyle = fmBorderStyleNone
   Me.WM.ForeColor = vbWhite
   Me.WM.value = ""
   If Me.SB.value = "" Then
    Me.SB.BorderStyle = fmBorderStyleSingle
    Me.SB.BorderColor = vbRed
    Exit Sub
    Else
    VSB = Replace(Me.SB.value, localDec, altDec)
    If vbaSep = "." Then
    VSB = Replace(VSB, altDec, Application.International(xlDecimalSeparator))
    End If
    VSB = Trim(VSB)
    If IsNumeric(VSB) Then
    SAB = CDbl(VSB)
    Else
    Me.SB.ForeColor = vbRed
    Exit Sub
    End If
    End If
    SAB = (SAB * 0.8)
    SAB = Format(SAB, "0")
    Me.WM.value = SAB
End If

End Sub
Private Sub CLCFIRENoInf()
Dim LibR As String

ER = 0
    Me.Strat2.BorderStyle = fmBorderStyleNone
    Me.Strat2.BorderColor = &HC0C0C0
    Me.LS2.BorderStyle = fmBorderStyleNone
    Me.LS2.BorderColor = &HC0C0C0
    
    If Me.Tf.value = "" Then
        Me.Tf.BorderStyle = fmBorderStyleSingle
        Me.Tf.BorderColor = vbRed
    ER = 1
    Else
    If IsNumeric(Me.Tf.value) Then
    FT = Me.Tf.value + 0
    If FT <> Int(FT) Then
    Me.Tf.ForeColor = vbRed
    ER = 1
    End If
    Else
    Me.Tf.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    
    If Me.Tc.value = "" Then
    Me.Tc.BorderStyle = fmBorderStyleSingle
    Me.Tc.BorderColor = vbRed
    ER = 1
    Else
    If IsNumeric(Me.Tc.value) Then
    CT = Me.Tc.value + 0
    If CT <> Int(CT) Then
    Me.Tc.ForeColor = vbRed
    ER = 1
    End If
    Else
    Me.Tc.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.Tr.value = "" Then
    Me.Tr.BorderStyle = fmBorderStyleSingle
    Me.Tr.BorderColor = vbRed
    ER = 1
    Else
    If IsNumeric(Me.Tr.value) Then
    RT = Me.Tr.value + 0
    If RT <> Int(RT) Then
    Me.Tr.ForeColor = vbRed
    ER = 1
    End If
    Else
    Me.Tr.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.SB.value = "" Then
    Me.SB.BorderStyle = fmBorderStyleSingle
    Me.SB.BorderColor = vbRed
    ER = 1
    Else
    VSAB = Replace(Me.SB.value, localDec, altDec)
    If vbaSep = "." Then
    VSAB = Replace(VSAB, altDec, Application.International(xlDecimalSeparator))
    End If
    VSAB = Trim(VSAB)
    If IsNumeric(VSAB) Then
    SAB = CDbl(VSAB)
    Else
    Me.SB.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.Montant.value = "" Then
    m = 0
    Else
    Vm = Replace(Me.Montant.value, localDec, altDec)
    If vbaSep = "." Then
    Vm = Replace(Vm, altDec, Application.International(xlDecimalSeparator))
    End If
    Vm = Trim(Vm)
    If IsNumeric(Vm) Then
    m = CDbl(Vm)
    Else
    Me.Montant.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.MEP.value = "" Then
    EPM = 0
    Else
    VEPM = Replace(Me.MEP.value, localDec, altDec)
    If vbaSep = "." Then
    VEPM = Replace(VEPM, altDec, Application.International(xlDecimalSeparator))
    End If
    VEPM = Trim(VEPM)
    If IsNumeric(VEPM) Then
    EPM = CDbl(VEPM)
    Else
    Me.MEP.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If m = 0 And EPM = 0 Then
    Me.Montant.BorderStyle = fmBorderStyleSingle
    Me.Montant.BorderColor = vbRed
    Me.MEP.BorderStyle = fmBorderStyleSingle
    Me.MEP.BorderColor = vbRed
    End If
    
    
    If Me.S1.value = "" Then
    Me.S1.BorderStyle = fmBorderStyleSingle
    Me.S1.BorderColor = vbRed
    ER = 1
    Else
    VS1 = Replace(Me.S1.value, localDec, altDec)
    If vbaSep = "." Then
    VS1 = Replace(VS1, altDec, Application.International(xlDecimalSeparator))
    End If
    VS1 = Trim(VS1)
    If IsNumeric(VS1) Then
    r1 = CDbl(VS1)
    Else
    Me.S1.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.S2.value = "" Then
    Me.S2.BorderStyle = fmBorderStyleSingle
    Me.S2.BorderColor = vbRed
    ER = 1
    Else
    VS2 = Replace(Me.S2.value, localDec, altDec)
    If vbaSep = "." Then
    VS2 = Replace(VS2, altDec, Application.International(xlDecimalSeparator))
    End If
    VS2 = Trim(VS2)
    If IsNumeric(VS2) Then
    r2 = CDbl(VS2)
    Else
    Me.S2.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.S3.value = "" Then
    Me.S3.BorderStyle = fmBorderStyleSingle
    Me.S3.BorderColor = vbRed
    ER = 1
    Else
    VS3 = Replace(Me.S3.value, localDec, altDec)
    If vbaSep = "." Then
    VS3 = Replace(VS3, altDec, Application.International(xlDecimalSeparator))
    End If
    VS3 = Trim(VS3)
    If IsNumeric(VS3) Then
    r3 = CDbl(VS3)
    Else
    Me.S3.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.S4.value = "" Then
    Me.S4.BorderStyle = fmBorderStyleSingle
    Me.S4.BorderColor = vbRed
    ER = 1
    Else
    VS4 = Replace(Me.S4.value, localDec, altDec)
    If vbaSep = "." Then
    VS4 = Replace(VS4, altDec, Application.International(xlDecimalSeparator))
    End If
    VS4 = Trim(VS4)
    If IsNumeric(VS4) Then
    r4 = CDbl(VS4)
    Else
    Me.S4.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.RIF.value = "" Then
    Me.RIF.BorderStyle = fmBorderStyleSingle
    Me.RIF.BorderColor = vbRed
    ER = 1
    Else
    VRIF = Replace(Me.RIF.value, localDec, altDec)
    If vbaSep = "." Then
    VRIF = Replace(VRIF, altDec, Application.International(xlDecimalSeparator))
    End If
    VRIF = Trim(VRIF)
    If IsNumeric(VRIF) Then
    YrY = CDbl(VRIF)
    Else
    Me.RIF.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.WM.value = "" Then
    Me.WM.BorderStyle = fmBorderStyleSingle
    Me.WM.BorderColor = vbRed
    ER = 1
    Else
    VWM = Replace(Me.WM.value, localDec, altDec)
    If vbaSep = "." Then
    VWM = Replace(VWM, altDec, Application.International(xlDecimalSeparator))
    End If
    VWM = Trim(VWM)
    If IsNumeric(VWM) Then
    Withdrawal = CDbl(VWM)
    Else
    Me.WM.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If ER = 1 Then
       Exit Sub
    End If
    

    r1 = (r1 / 100)
    r2 = (r2 / 100)
    r3 = (r3 / 100)
    r4 = (r4 / 100)
    YrY = (YrY / 100)
    
    n = 1
    CheckYear = 1
    
    
    Dim MonthlyGrowthAccum As Double, MonthlyGrowthDist As Double, MonthlyInflation As Double
    r1 = (1 + r1) ^ (1 / 12) - 1
    r2 = (1 + r2) ^ (1 / 12) - 1
    r3 = (1 + r3) ^ (1 / 12) - 1
    r4 = (1 + r4) ^ (1 / 12) - 1
    rinf = (1 + YrY) ^ (1 / 12) - 1
    r11 = r1 - rinf
    r21 = r2 - rinf
    r31 = r3 - rinf
    r41 = r4 - rinf
    YrY = (1 + YrY)

    ' Convert ages to months
    Dim MonthsAccumulation As Long, MonthsDistribution As Long
    MonthsAccumulation = (RT - CT) * 12
    MonthsDistribution = (FT - RT) * 12

    ' Phase 1: Accumulation Phase
    Dim P As Double, MonthlyContribution As Double
    FIRE_Number1 = m
    FIRE_Number2 = m
    FIRE_Number3 = m

    ' Start with the current accumulated amount
    Withdrawal = Withdrawal / 12
     
    For i = 1 To MonthsAccumulation
        
        FIRE_Number1 = (FIRE_Number1 + EPM) * (1 + r11) ' Monthly growth with contributions
        FIRE_Number2 = (FIRE_Number2 + EPM) * (1 + r21) ' Monthly growth with contributions
        FIRE_Number3 = (FIRE_Number3 + EPM) * (1 + r31) ' Monthly growth with contributions
        If CheckYear = 12 Then
           n = n + 1
           CheckYear = 0
        End If
        CheckYear = CheckYear + 1
    Next i

    ' Phase 2: Find Required FIRE Number at Retirement

    ' Start with an estimate equal to accumulated capital
    P1 = FIRE_Number1
    P2 = FIRE_Number2
    P3 = FIRE_Number3

    ' Iterative approach: Adjust FIRENumber until balance reaches ~0 at EndAge
  
        MissingYears1 = 0
        MissingYears2 = 0
        MissingYears3 = 0
        CheckYear = 1
        ' Simulate withdrawals until EndAge
        For i = 1 To MonthsDistribution

            If P1 >= 0 Then
            P1 = (P1 - Withdrawal) * (1 + r41)
            End If
            If P2 >= 0 Then
            P2 = (P2 - Withdrawal) * (1 + r41)
            End If
            If P3 >= 0 Then
            P3 = (P3 - Withdrawal) * (1 + r41)
            End If
            
        If CheckYear = 12 Then
           n = n + 1
           CheckYear = 0
           
           If P1 < 0 Then
              MissingYears1 = MissingYears1 + 1
           End If
           If P2 < 0 Then
              MissingYears2 = MissingYears2 + 1
           End If
           If P3 < 0 Then
              MissingYears3 = MissingYears3 + 1
           End If
           
        End If
        CheckYear = CheckYear + 1
        Next i

    FIRE_Number1 = Format(FIRE_Number1, "# ### ### ### ##0")
    FIRE_Number2 = Format(FIRE_Number2, "# ### ### ### ##0")
    FIRE_Number3 = Format(FIRE_Number3, "# ### ### ### ##0")
Me.LFDD1.ForeColor = vbWhite
Me.LFDD2.ForeColor = vbWhite
Me.LFDD3.ForeColor = vbWhite
Me.LF1.Caption = FIRE_Number1
Me.LF2.Caption = FIRE_Number2
Me.LF3.Caption = FIRE_Number3
Me.LA1.Caption = MissingYears1
Me.LA2.Caption = MissingYears2
Me.LA3.Caption = MissingYears3
Me.LAA1.ForeColor = vbWhite
Me.LAA2.ForeColor = vbWhite
Me.LAA3.ForeColor = vbWhite
End Sub


Private Sub CLCFIRE_Click()
If CheckInf = True Then
CLCFIREInf
ElseIf CheckInf = False Then
CLCFIRENoInf
End If
End Sub
Private Sub CLCFIREInf()
Dim LibR As String

ER = 0
    Me.Strat2.BorderStyle = fmBorderStyleNone
    Me.Strat2.BorderColor = &HC0C0C0
    Me.LS2.BorderStyle = fmBorderStyleNone
    Me.LS2.BorderColor = &HC0C0C0
    
    If Me.Tf.value = "" Then
        Me.Tf.BorderStyle = fmBorderStyleSingle
        Me.Tf.BorderColor = vbRed
    ER = 1
    Else
    If IsNumeric(Me.Tf.value) Then
    FT = Me.Tf.value + 0
    If FT <> Int(FT) Then
    Me.Tf.ForeColor = vbRed
    ER = 1
    End If
    Else
    Me.Tf.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    
    If Me.Tc.value = "" Then
    Me.Tc.BorderStyle = fmBorderStyleSingle
    Me.Tc.BorderColor = vbRed
    ER = 1
    Else
    If IsNumeric(Me.Tc.value) Then
    CT = Me.Tc.value + 0
    If CT <> Int(CT) Then
    Me.Tc.ForeColor = vbRed
    ER = 1
    End If
    Else
    Me.Tc.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.Tr.value = "" Then
    Me.Tr.BorderStyle = fmBorderStyleSingle
    Me.Tr.BorderColor = vbRed
    ER = 1
    Else
    If IsNumeric(Me.Tr.value) Then
    RT = Me.Tr.value + 0
    If RT <> Int(RT) Then
    Me.Tr.ForeColor = vbRed
    ER = 1
    End If
    Else
    Me.Tr.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.SB.value = "" Then
    Me.SB.BorderStyle = fmBorderStyleSingle
    Me.SB.BorderColor = vbRed
    ER = 1
    Else
    VSAB = Replace(Me.SB.value, localDec, altDec)
    If vbaSep = "." Then
    VSAB = Replace(VSAB, altDec, Application.International(xlDecimalSeparator))
    End If
    VSAB = Trim(VSAB)
    If IsNumeric(VSAB) Then
    SAB = CDbl(VSAB)
    Else
    Me.SB.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.Montant.value = "" Then
    m = 0
    Else
    Vm = Replace(Me.Montant.value, localDec, altDec)
    If vbaSep = "." Then
    Vm = Replace(Vm, altDec, Application.International(xlDecimalSeparator))
    End If
    Vm = Trim(Vm)
    If IsNumeric(Vm) Then
    m = CDbl(Vm)
    Else
    Me.Montant.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.MEP.value = "" Then
    EPM = 0
    Else
    VEPM = Replace(Me.MEP.value, localDec, altDec)
    If vbaSep = "." Then
    VEPM = Replace(VEPM, altDec, Application.International(xlDecimalSeparator))
    End If
    VEPM = Trim(VEPM)
    If IsNumeric(VEPM) Then
    EPM = CDbl(VEPM)
    Else
    Me.MEP.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If m = 0 And EPM = 0 Then
    Me.Montant.BorderStyle = fmBorderStyleSingle
    Me.Montant.BorderColor = vbRed
    Me.MEP.BorderStyle = fmBorderStyleSingle
    Me.MEP.BorderColor = vbRed
    End If
    
    
    If Me.S1.value = "" Then
    Me.S1.BorderStyle = fmBorderStyleSingle
    Me.S1.BorderColor = vbRed
    ER = 1
    Else
    VS1 = Replace(Me.S1.value, localDec, altDec)
    If vbaSep = "." Then
    VS1 = Replace(VS1, altDec, Application.International(xlDecimalSeparator))
    End If
    VS1 = Trim(VS1)
    If IsNumeric(VS1) Then
    r1 = CDbl(VS1)
    Else
    Me.S1.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.S2.value = "" Then
    Me.S2.BorderStyle = fmBorderStyleSingle
    Me.S2.BorderColor = vbRed
    ER = 1
    Else
    VS2 = Replace(Me.S2.value, localDec, altDec)
    If vbaSep = "." Then
    VS2 = Replace(VS2, altDec, Application.International(xlDecimalSeparator))
    End If
    VS2 = Trim(VS2)
    If IsNumeric(VS2) Then
    r2 = CDbl(VS2)
    Else
    Me.S2.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.S3.value = "" Then
    Me.S3.BorderStyle = fmBorderStyleSingle
    Me.S3.BorderColor = vbRed
    ER = 1
    Else
    VS3 = Replace(Me.S3.value, localDec, altDec)
    If vbaSep = "." Then
    VS3 = Replace(VS3, altDec, Application.International(xlDecimalSeparator))
    End If
    VS3 = Trim(VS3)
    If IsNumeric(VS3) Then
    r3 = CDbl(VS3)
    Else
    Me.S3.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.S4.value = "" Then
    Me.S4.BorderStyle = fmBorderStyleSingle
    Me.S4.BorderColor = vbRed
    ER = 1
    Else
    VS4 = Replace(Me.S4.value, localDec, altDec)
    If vbaSep = "." Then
    VS4 = Replace(VS4, altDec, Application.International(xlDecimalSeparator))
    End If
    VS4 = Trim(VS4)
    If IsNumeric(VS4) Then
    r4 = CDbl(VS4)
    Else
    Me.S4.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.RIF.value = "" Then
    Me.RIF.BorderStyle = fmBorderStyleSingle
    Me.RIF.BorderColor = vbRed
    ER = 1
    Else
    VRIF = Replace(Me.RIF.value, localDec, altDec)
    If vbaSep = "." Then
    VRIF = Replace(VRIF, altDec, Application.International(xlDecimalSeparator))
    End If
    VRIF = Trim(VRIF)
    If IsNumeric(VRIF) Then
    YrY = CDbl(VRIF)
    Else
    Me.RIF.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.WM.value = "" Then
    Me.WM.BorderStyle = fmBorderStyleSingle
    Me.WM.BorderColor = vbRed
    ER = 1
    Else
    VWM = Replace(Me.WM.value, localDec, altDec)
    If vbaSep = "." Then
    VWM = Replace(VWM, altDec, Application.International(xlDecimalSeparator))
    End If
    VWM = Trim(VWM)
    If IsNumeric(VWM) Then
    Withdrawal = CDbl(VWM)
    Else
    Me.WM.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If ER = 1 Then
       Exit Sub
    End If
    

    r1 = (r1 / 100)
    r2 = (r2 / 100)
    r3 = (r3 / 100)
    r4 = (r4 / 100)
    YrY = (YrY / 100)
    n = 1
    CheckYear = 1
    
    
    Dim MonthlyGrowthAccum As Double, MonthlyGrowthDist As Double, MonthlyInflation As Double
    r1 = (1 + r1) ^ (1 / 12) - 1
    r2 = (1 + r2) ^ (1 / 12) - 1
    r3 = (1 + r3) ^ (1 / 12) - 1
    r4 = (1 + r4) ^ (1 / 12) - 1
    YrY = (1 + YrY)

    ' Convert ages to months
    Dim MonthsAccumulation As Long, MonthsDistribution As Long
    MonthsAccumulation = (RT - CT) * 12
    MonthsDistribution = (FT - RT) * 12

    ' Phase 1: Accumulation Phase
    Dim P As Double, MonthlyContribution As Double
    FIRE_Number1 = m
    FIRE_Number2 = m
    FIRE_Number3 = m

    ' Start with the current accumulated amount
    Withdrawal = Withdrawal / 12
     
    For i = 1 To MonthsAccumulation
        
        FIRE_Number1 = (FIRE_Number1 + EPM) * (1 + r1) ' Monthly growth with contributions
        FIRE_Number2 = (FIRE_Number2 + EPM) * (1 + r2) ' Monthly growth with contributions
        FIRE_Number3 = (FIRE_Number3 + EPM) * (1 + r3) ' Monthly growth with contributions
        If CheckYear = 12 Then
           n = n + 1
           EPM = EPM * YrY
           Withdrawal = Withdrawal * YrY
           CheckYear = 0
        End If
        CheckYear = CheckYear + 1
    Next i

    ' Phase 2: Find Required FIRE Number at Retirement

    ' Start with an estimate equal to accumulated capital
    P1 = FIRE_Number1
    P2 = FIRE_Number2
    P3 = FIRE_Number3

    ' Iterative approach: Adjust FIRENumber until balance reaches ~0 at EndAge
  
        MissingYears1 = 0
        MissingYears2 = 0
        MissingYears3 = 0
        CheckYear = 1
        ' Simulate withdrawals until EndAge
        For i = 1 To MonthsDistribution

            If P1 >= 0 Then
            P1 = (P1 - Withdrawal) * (1 + r4)
            End If
            If P2 >= 0 Then
            P2 = (P2 - Withdrawal) * (1 + r4)
            End If
            If P3 >= 0 Then
            P3 = (P3 - Withdrawal) * (1 + r4)
            End If
            
        If CheckYear = 12 Then
           n = n + 1
           Withdrawal = Withdrawal * YrY
           CheckYear = 0
           
           If P1 < 0 Then
              MissingYears1 = MissingYears1 + 1
           End If
           If P2 < 0 Then
              MissingYears2 = MissingYears2 + 1
           End If
           If P3 < 0 Then
              MissingYears3 = MissingYears3 + 1
           End If
           
        End If
        CheckYear = CheckYear + 1
        Next i

       
    
    FIRE_Number1 = Format(FIRE_Number1, "# ### ### ### ##0")
    FIRE_Number2 = Format(FIRE_Number2, "# ### ### ### ##0")
    FIRE_Number3 = Format(FIRE_Number3, "# ### ### ### ##0")
Me.LFDD1.ForeColor = vbWhite
Me.LFDD2.ForeColor = vbWhite
Me.LFDD3.ForeColor = vbWhite
Me.LF1.Caption = FIRE_Number1
Me.LF2.Caption = FIRE_Number2
Me.LF3.Caption = FIRE_Number3
Me.LA1.Caption = MissingYears1
Me.LA2.Caption = MissingYears2
Me.LA3.Caption = MissingYears3
Me.LAA1.ForeColor = vbWhite
Me.LAA2.ForeColor = vbWhite
Me.LAA3.ForeColor = vbWhite
End Sub

Private Sub Label44_Click()

End Sub

Private Sub Label45_Click()

End Sub

Private Sub LRIF_Click()

End Sub

Private Sub LS2_Click()

End Sub

Private Sub MEP_Change()
Me.MEP.BorderStyle = fmBorderStyleNone
Me.MEP.ForeColor = vbWhite
If Q = 1 Then
Q = 0
Else
Me.QMEP.value = False
End If
End Sub

Private Sub Montant_Change()
Me.Montant.BorderStyle = fmBorderStyleNone
Me.Montant.ForeColor = vbWhite
End Sub

Private Sub QMEPNoinf()
ER = 0
    If Me.QMEP.value = False Then
    Me.Strat2.BorderStyle = fmBorderStyleNone
    Me.Strat2.BorderColor = &HC0C0C0
    Me.LS2.BorderStyle = fmBorderStyleNone
    Me.LS2.BorderColor = &HC0C0C0
       ER = 1
    End If
    If Me.Tf.value = "" Then
        Me.Tf.BorderStyle = fmBorderStyleSingle
        Me.Tf.BorderColor = vbRed
    ER = 1
    Else
    If IsNumeric(Me.Tf.value) Then
    FT = Me.Tf.value + 0
    If FT <> Int(FT) Then
    Me.Tf.ForeColor = vbRed
    ER = 1
    End If
    Else
    Me.Tf.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    
    If Me.Tc.value = "" Then
    Me.Tc.BorderStyle = fmBorderStyleSingle
    Me.Tc.BorderColor = vbRed
    ER = 1
    Else
    If IsNumeric(Me.Tc.value) Then
    CT = Me.Tc.value + 0
    If CT <> Int(CT) Then
    Me.Tc.ForeColor = vbRed
    ER = 1
    End If
    Else
    Me.Tc.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.Tr.value = "" Then
    Me.Tr.BorderStyle = fmBorderStyleSingle
    Me.Tr.BorderColor = vbRed
    ER = 1
    Else
    If IsNumeric(Me.Tr.value) Then
    RT = Me.Tr.value + 0
    If RT <> Int(RT) Then
    Me.Tr.ForeColor = vbRed
    ER = 1
    End If
    Else
    Me.Tr.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.SB.value = "" Then
    Me.SB.BorderStyle = fmBorderStyleSingle
    Me.SB.BorderColor = vbRed
    ER = 1
    Else
    VSB = Replace(Me.SB.value, localDec, altDec)
    If vbaSep = "." Then
    VSB = Replace(VSB, altDec, Application.International(xlDecimalSeparator))
    End If
    VSB = Trim(VSB)
    If IsNumeric(VSB) Then
    SAB = CDbl(VSB)
    Else
    Me.SB.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.Montant.value = "" Then
    m = 0
    Else
    Vm = Replace(Me.Montant.value, localDec, altDec)
    If vbaSep = "." Then
    Vm = Replace(Vm, altDec, Application.International(xlDecimalSeparator))
    End If
    Vm = Trim(Vm)
    If IsNumeric(Vm) Then
    m = CDbl(Vm)
    Else
    Me.Montant.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.S2.value = "" Then
    Me.S2.BorderStyle = fmBorderStyleSingle
    Me.S2.BorderColor = vbRed
    ER = 1
    Else
    VS2 = Replace(Me.S2.value, localDec, altDec)
    If vbaSep = "." Then
    VS2 = Replace(VS2, altDec, Application.International(xlDecimalSeparator))
    End If
    VS2 = Trim(VS2)
    If IsNumeric(VS2) Then
    r2 = CDbl(VS2)
    Else
    Me.S2.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.S4.value = "" Then
    Me.S4.BorderStyle = fmBorderStyleSingle
    Me.S4.BorderColor = vbRed
    ER = 1
    Else
    VS4 = Replace(Me.S4.value, localDec, altDec)
    If vbaSep = "." Then
    VS4 = Replace(VS4, altDec, Application.International(xlDecimalSeparator))
    End If
    VS4 = Trim(VS4)
    If IsNumeric(VS4) Then
    r4 = CDbl(VS4)
    Else
    Me.S4.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.RIF.value = "" Then
    Me.RIF.BorderStyle = fmBorderStyleSingle
    Me.RIF.BorderColor = vbRed
    ER = 1
    Else
    VRIF = Replace(Me.RIF.value, localDec, altDec)
    If vbaSep = "." Then
    VRIF = Replace(VRIF, altDec, Application.International(xlDecimalSeparator))
    End If
    VRIF = Trim(VRIF)
    If IsNumeric(VRIF) Then
    YrY = CDbl(VRIF)
    Else
    Me.RIF.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.WM.value = "" Then
    Me.WM.BorderStyle = fmBorderStyleSingle
    Me.WM.BorderColor = vbRed
    ER = 1
    Else
    VWM = Replace(Me.WM.value, localDec, altDec)
    If vbaSep = "." Then
    VWM = Replace(VWM, altDec, Application.International(xlDecimalSeparator))
    End If
    VWM = Trim(VWM)
    If IsNumeric(VWM) Then
    Withdrawal = CDbl(VWM)
    Else
    Me.WM.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If ER = 1 Then
      Me.QMEP.value = False
      Exit Sub
    End If
    
    r2 = (r2 / 100)
    r4 = (r4 / 100)
    YrY = (YrY / 100)
    
    Tolerance = 100              ' Convergence tolerance for the monthly deposit calculation
    MaxIterations = 10000          ' Maximum number of iterations allowed
    
    Dim MonthlyGrowthAccum As Double, MonthlyGrowthDist As Double, MonthlyInflation As Double
    r2 = (1 + r2) ^ (1 / 12) - 1
    r4 = (1 + r4) ^ (1 / 12) - 1
    rinf = (1 + YrY) ^ (1 / 12) - 1
    r5 = r2 - rinf
    r6 = r4 - rinf
    YrY = (1 + YrY)

    ' Convert ages to months
    Dim MonthsAccumulation As Long, MonthsDistribution As Long
    MonthsAccumulation = (RT - CT) * 12
    MonthsDistribution = (FT - RT) * 12

    ' Phase 1: Accumulation Phase
    Dim P As Double, MonthlyContribution As Double

    
    MonthlyDeposit = 100 ' Start with an initial guess for the monthly deposit
    IterationCount = 0
    Do
    Me.MousePointer = fmMousePointerHourGlass
        n = 1
        CheckYear = 1
        EPM = MonthlyDeposit
        IterationCount = IterationCount + 1
        FIRE_Number2 = m
        CurrentWithdrawal = Withdrawal / 12
     
    For i = 1 To MonthsAccumulation
    
        FIRE_Number2 = (FIRE_Number2 + EPM) * (1 + r5) ' Monthly growth with contributions
        If CheckYear = 12 Then
           n = n + 1
           CheckYear = 0
        End If
        CheckYear = CheckYear + 1
    
    Next i

    ' Phase 2: Find Required FIRE Number at Retirement

    ' Start with an estimate equal to accumulated capital
    P2 = FIRE_Number2

    ' Iterative approach: Adjust FIRENumber until balance reaches ~0 at EndAge
  
        CheckYear = 1
    For i = 1 To MonthsDistribution
        ' Simulate withdrawals until EndAge
            P2 = (P2 - CurrentWithdrawal) * (1 + r6)
            
        If CheckYear = 12 Then
           n = n + 1
           CheckYear = 0
           
        End If
        CheckYear = CheckYear + 1
        
        Next i
 
        If P2 >= 0 Then Exit Do
        


        ' Adjust the monthly deposit based on the difference
        If P2 < FIRE_Number2 Then
            MonthlyDeposit = MonthlyDeposit * 1.01 ' Increase deposit if not enough
        Else
            MonthlyDeposit = MonthlyDeposit * 0.99 ' Decrease deposit if too much
        End If
    Loop While IterationCount < MaxIterations
        
    Me.MousePointer = fmMousePointerDefault
    ' Output results
    If IterationCount >= MaxIterations Then
        MsgBox "Calculation did not converge. Try increasing the maximum iterations or adjusting the tolerance.", vbExclamation
        Me.MEP.value = ""
        Me.QMEP.value = False
    Else
        Me.LFDD1.ForeColor = vbBlack
        Me.LFDD2.ForeColor = vbWhite
        Me.Strat2.BorderStyle = fmBorderStyleSingle
        Me.Strat2.BorderColor = vbWhite
        Me.LS2.BorderStyle = fmBorderStyleSingle
        Me.LS2.BorderColor = vbWhite
        Me.LFDD3.ForeColor = vbBlack
        Me.LF1.Caption = ""
        Me.LF2.Caption = Format(FIRE_Number2, "# ### ### ### ##0")
        Me.LF3.Caption = ""
        Me.LA1.Caption = ""
        Me.LA2.Caption = ""
        Me.LA3.Caption = ""
        Me.LAA1.ForeColor = vbBlack
        Me.LAA2.ForeColor = vbBlack
        Me.LAA3.ForeColor = vbBlack
    End If
    Q = 1
    Me.MEP.value = Format(MonthlyDeposit, "0")
End Sub


Private Sub QMEP_Click()
If CheckInf = True Then
QMEPInf
ElseIf CheckInf = False Then
QMEPNoinf
End If
End Sub
Private Sub QMEPInf()
ER = 0
    If Me.QMEP.value = False Then
    Me.Strat2.BorderStyle = fmBorderStyleNone
    Me.Strat2.BorderColor = &HC0C0C0
    Me.LS2.BorderStyle = fmBorderStyleNone
    Me.LS2.BorderColor = &HC0C0C0
       ER = 1
    End If
    If Me.Tf.value = "" Then
        Me.Tf.BorderStyle = fmBorderStyleSingle
        Me.Tf.BorderColor = vbRed
    ER = 1
    Else
    If IsNumeric(Me.Tf.value) Then
    FT = Me.Tf.value + 0
    If FT <> Int(FT) Then
    Me.Tf.ForeColor = vbRed
    ER = 1
    End If
    Else
    Me.Tf.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    
    If Me.Tc.value = "" Then
    Me.Tc.BorderStyle = fmBorderStyleSingle
    Me.Tc.BorderColor = vbRed
    ER = 1
    Else
    If IsNumeric(Me.Tc.value) Then
    CT = Me.Tc.value + 0
    If CT <> Int(CT) Then
    Me.Tc.ForeColor = vbRed
    ER = 1
    End If
    Else
    Me.Tc.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.Tr.value = "" Then
    Me.Tr.BorderStyle = fmBorderStyleSingle
    Me.Tr.BorderColor = vbRed
    ER = 1
    Else
    If IsNumeric(Me.Tr.value) Then
    RT = Me.Tr.value + 0
    If RT <> Int(RT) Then
    Me.Tr.ForeColor = vbRed
    ER = 1
    End If
    Else
    Me.Tr.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.SB.value = "" Then
    Me.SB.BorderStyle = fmBorderStyleSingle
    Me.SB.BorderColor = vbRed
    ER = 1
    Else
    VSB = Replace(Me.SB.value, localDec, altDec)
    If vbaSep = "." Then
    VSB = Replace(VSB, altDec, Application.International(xlDecimalSeparator))
    End If
    VSB = Trim(VSB)
    If IsNumeric(VSB) Then
    SAB = CDbl(VSB)
    Else
    Me.SB.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.Montant.value = "" Then
    m = 0
    Else
    Vm = Replace(Me.Montant.value, localDec, altDec)
    If vbaSep = "." Then
    Vm = Replace(Vm, altDec, Application.International(xlDecimalSeparator))
    End If
    Vm = Trim(Vm)
    If IsNumeric(Vm) Then
    m = CDbl(Vm)
    Else
    Me.Montant.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.S2.value = "" Then
    Me.S2.BorderStyle = fmBorderStyleSingle
    Me.S2.BorderColor = vbRed
    ER = 1
    Else
    VS2 = Replace(Me.S2.value, localDec, altDec)
    If vbaSep = "." Then
    VS2 = Replace(VS2, altDec, Application.International(xlDecimalSeparator))
    End If
    VS2 = Trim(VS2)
    If IsNumeric(VS2) Then
    r2 = CDbl(VS2)
    Else
    Me.S2.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.S4.value = "" Then
    Me.S4.BorderStyle = fmBorderStyleSingle
    Me.S4.BorderColor = vbRed
    ER = 1
    Else
    VS4 = Replace(Me.S4.value, localDec, altDec)
    If vbaSep = "." Then
    VS4 = Replace(VS4, altDec, Application.International(xlDecimalSeparator))
    End If
    VS4 = Trim(VS4)
    If IsNumeric(VS4) Then
    r4 = CDbl(VS4)
    Else
    Me.S4.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.RIF.value = "" Then
    Me.RIF.BorderStyle = fmBorderStyleSingle
    Me.RIF.BorderColor = vbRed
    ER = 1
    Else
    VRIF = Replace(Me.RIF.value, localDec, altDec)
    If vbaSep = "." Then
    VRIF = Replace(VRIF, altDec, Application.International(xlDecimalSeparator))
    End If
    VRIF = Trim(VRIF)
    If IsNumeric(VRIF) Then
    YrY = CDbl(VRIF)
    Else
    Me.RIF.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.WM.value = "" Then
    Me.WM.BorderStyle = fmBorderStyleSingle
    Me.WM.BorderColor = vbRed
    ER = 1
    Else
    VWM = Replace(Me.WM.value, localDec, altDec)
    If vbaSep = "." Then
    VWM = Replace(VWM, altDec, Application.International(xlDecimalSeparator))
    End If
    VWM = Trim(VWM)
    If IsNumeric(VWM) Then
    Withdrawal = CDbl(VWM)
    Else
    Me.WM.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If ER = 1 Then
      Me.QMEP.value = False
      Exit Sub
    End If
    
    
    r2 = (r2 / 100)
    r4 = (r4 / 100)
    YrY = (YrY / 100)
    Tolerance = 100              ' Convergence tolerance for the monthly deposit calculation
    MaxIterations = 10000          ' Maximum number of iterations allowed
    
    Dim MonthlyGrowthAccum As Double, MonthlyGrowthDist As Double, MonthlyInflation As Double
    r2 = (1 + r2) ^ (1 / 12) - 1
    r4 = (1 + r4) ^ (1 / 12) - 1
    YrY = (1 + YrY)

    ' Convert ages to months
    Dim MonthsAccumulation As Long, MonthsDistribution As Long
    MonthsAccumulation = (RT - CT) * 12
    MonthsDistribution = (FT - RT) * 12

    ' Phase 1: Accumulation Phase
    Dim P As Double, MonthlyContribution As Double

    
    MonthlyDeposit = 100 ' Start with an initial guess for the monthly deposit
    IterationCount = 0
    Do
    Me.MousePointer = fmMousePointerHourGlass
        n = 1
        CheckYear = 1
        EPM = MonthlyDeposit
        IterationCount = IterationCount + 1
        FIRE_Number2 = m
        CurrentWithdrawal = Withdrawal / 12
     
    For i = 1 To MonthsAccumulation
    
        FIRE_Number2 = (FIRE_Number2 + EPM) * (1 + r2) ' Monthly growth with contributions
        If CheckYear = 12 Then
           n = n + 1
           EPM = EPM * YrY
           CurrentWithdrawal = CurrentWithdrawal * YrY
           CheckYear = 0
        End If
        CheckYear = CheckYear + 1
    
    Next i

    ' Phase 2: Find Required FIRE Number at Retirement

    ' Start with an estimate equal to accumulated capital
    P2 = FIRE_Number2

    ' Iterative approach: Adjust FIRENumber until balance reaches ~0 at EndAge
  
        CheckYear = 1
    For i = 1 To MonthsDistribution
        ' Simulate withdrawals until EndAge
            P2 = (P2 - CurrentWithdrawal) * (1 + r4)
            
        If CheckYear = 12 Then
           n = n + 1
           CurrentWithdrawal = CurrentWithdrawal * YrY
           CheckYear = 0
           
        End If
        CheckYear = CheckYear + 1
        
        Next i
 
        If P2 >= 0 Then Exit Do
        


        ' Adjust the monthly deposit based on the difference
        If P2 < FIRE_Number2 Then
            MonthlyDeposit = MonthlyDeposit * 1.01 ' Increase deposit if not enough
        Else
            MonthlyDeposit = MonthlyDeposit * 0.99 ' Decrease deposit if too much
        End If
    Loop While IterationCount < MaxIterations
        
    Me.MousePointer = fmMousePointerDefault
    ' Output results
    If IterationCount >= MaxIterations Then
        MsgBox "Calculation did not converge. Try increasing the maximum iterations or adjusting the tolerance.", vbExclamation
        Me.MEP.value = ""
        Me.QMEP.value = False
    Else
        Me.LFDD1.ForeColor = vbBlack
        Me.LFDD2.ForeColor = vbWhite
        Me.Strat2.BorderStyle = fmBorderStyleSingle
        Me.Strat2.BorderColor = vbWhite
        Me.LS2.BorderStyle = fmBorderStyleSingle
        Me.LS2.BorderColor = vbWhite
        Me.LFDD3.ForeColor = vbBlack
        Me.LF1.Caption = ""
        Me.LF2.Caption = Format(FIRE_Number2, "# ### ### ### ##0")
        Me.LF3.Caption = ""
        Me.LA1.Caption = ""
        Me.LA2.Caption = ""
        Me.LA3.Caption = ""
        Me.LAA1.ForeColor = vbBlack
        Me.LAA2.ForeColor = vbBlack
        Me.LAA3.ForeColor = vbBlack
    End If
    Q = 1
    Me.MEP.value = Format(MonthlyDeposit, "0")
    
End Sub

Private Sub Quitt_Click()
Unload Me
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub RIF_Change()
Me.RIF.BorderStyle = fmBorderStyleNone
Me.RIF.ForeColor = vbWhite
End Sub

Private Sub S1_Change()
Me.S1.BorderStyle = fmBorderStyleNone
Me.S1.ForeColor = vbWhite
End Sub

Private Sub S2_Change()
Me.S2.BorderStyle = fmBorderStyleNone
Me.S2.ForeColor = vbWhite
End Sub

Private Sub S3_Change()
Me.S3.BorderStyle = fmBorderStyleNone
Me.S3.ForeColor = vbWhite
End Sub

Private Sub S4_Change()
Me.S4.BorderStyle = fmBorderStyleNone
Me.S4.ForeColor = vbWhite
End Sub

Private Sub SB_Change()
Me.CHKWM.value = False
Me.SB.BorderStyle = fmBorderStyleNone
Me.SB.ForeColor = vbWhite
End Sub


Private Sub Tc_Change()
Me.Tc.BorderStyle = fmBorderStyleNone
Me.Tc.ForeColor = vbWhite
End Sub

Private Sub Tf_Change()
Me.Tf.BorderStyle = fmBorderStyleNone
Me.Tf.ForeColor = vbWhite
End Sub

Private Sub Tr_Change()
Me.Tr.BorderStyle = fmBorderStyleNone
Me.Tr.ForeColor = vbWhite
End Sub
Private Sub ValiderNoInf()

Dim LibR As String
ER = 0
    Me.Strat2.BorderStyle = fmBorderStyleNone
    Me.Strat2.BorderColor = &HC0C0C0
    Me.LS2.BorderStyle = fmBorderStyleNone
    Me.LS2.BorderColor = &HC0C0C0
    
    If Me.Tf.value = "" Then
        Me.Tf.BorderStyle = fmBorderStyleSingle
        Me.Tf.BorderColor = vbRed
    ER = 1
    Else
    If IsNumeric(Me.Tf.value) Then
    FT = Me.Tf.value + 0
    If FT <> Int(FT) Then
    Me.Tf.ForeColor = vbRed
    ER = 1
    End If
    If ER = 0 Then
    TbWEP.ListColumns("tf").DataBodyRange(1, 1).value = FT
    TbWEP.ListColumns("tf").DataBodyRange(2, 1).value = FT
    TbWEP.ListColumns("tf").DataBodyRange(3, 1).value = FT
    End If
    Else
    Me.Tf.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    
    If Me.Tc.value = "" Then
    Me.Tc.BorderStyle = fmBorderStyleSingle
    Me.Tc.BorderColor = vbRed
    ER = 1
    Else
    If IsNumeric(Me.Tc.value) Then
    CT = Me.Tc.value + 0
    If CT <> Int(CT) Then
    Me.Tc.ForeColor = vbRed
    ER = 1
    End If
    If ER = 0 Then
    TbWEP.ListColumns("tc").DataBodyRange(1, 1).value = CT
    TbWEP.ListColumns("tc").DataBodyRange(2, 1).value = CT
    TbWEP.ListColumns("tc").DataBodyRange(3, 1).value = CT
    End If
    Else
    Me.Tc.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.Tr.value = "" Then
    Me.Tr.BorderStyle = fmBorderStyleSingle
    Me.Tr.BorderColor = vbRed
    ER = 1
    Else
    If IsNumeric(Me.Tr.value) Then
    RT = Me.Tr.value + 0
    If RT <> Int(RT) Then
    Me.Tr.ForeColor = vbRed
    ER = 1
    End If
    If ER = 0 Then
    TbWEP.ListColumns("tr").DataBodyRange(1, 1).value = RT
    TbWEP.ListColumns("tr").DataBodyRange(2, 1).value = RT
    TbWEP.ListColumns("tr").DataBodyRange(3, 1).value = RT
    End If
    Else
    Me.Tr.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.SB.value = "" Then
    Me.SB.BorderStyle = fmBorderStyleSingle
    Me.SB.BorderColor = vbRed
    ER = 1
    Else
    VSB = Replace(Me.SB.value, localDec, altDec)
    If vbaSep = "." Then
    VSB = Replace(VSB, altDec, Application.International(xlDecimalSeparator))
    End If
    VSB = Trim(VSB)
    If IsNumeric(VSB) Then
    SAB = CDbl(VSB)
    If ER = 0 Then
    TbWEP.ListColumns("SB").DataBodyRange(1, 1).value = SAB
    TbWEP.ListColumns("SB").DataBodyRange(2, 1).value = SAB
    TbWEP.ListColumns("SB").DataBodyRange(3, 1).value = SAB
    End If
    Else
    Me.SB.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.Montant.value = "" Then
    m = 0
    Else
    Vm = Replace(Me.Montant.value, localDec, altDec)
    If vbaSep = "." Then
    Vm = Replace(Vm, altDec, Application.International(xlDecimalSeparator))
    End If
    Vm = Trim(Vm)
    If IsNumeric(Vm) Then
    m = CDbl(Vm)
    If ER = 0 Then
    TbWEP.ListColumns("M").DataBodyRange(1, 1).value = m
    TbWEP.ListColumns("M").DataBodyRange(2, 1).value = m
    TbWEP.ListColumns("M").DataBodyRange(3, 1).value = m
    End If
    Else
    Me.Montant.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.MEP.value = "" Then
    EPM = 0
    Else
    VMEPM = Replace(Me.MEP.value, localDec, altDec)
    If vbaSep = "." Then
    VMEPM = Replace(VMEPM, altDec, Application.International(xlDecimalSeparator))
    End If
    VMEPM = Trim(VMEPM)
    If IsNumeric(VMEPM) Then
    EPM = CDbl(VEPM)
    If ER = 0 Then
    TbWEP.ListColumns("MEP").DataBodyRange(1, 1).value = EPM
    TbWEP.ListColumns("MEP").DataBodyRange(2, 1).value = EPM
    TbWEP.ListColumns("MEP").DataBodyRange(3, 1).value = EPM
    End If
    Else
    Me.MEP.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If m = 0 And EPM = 0 Then
    Me.Montant.BorderStyle = fmBorderStyleSingle
    Me.Montant.BorderColor = vbRed
    Me.MEP.BorderStyle = fmBorderStyleSingle
    Me.MEP.BorderColor = vbRed
    End If
    
    
    If Me.S1.value = "" Then
    Me.S1.BorderStyle = fmBorderStyleSingle
    Me.S1.BorderColor = vbRed
    ER = 1
    Else
    VS1 = Replace(Me.S1.value, localDec, altDec)
    If vbaSep = "." Then
    VS1 = Replace(VS1, altDec, Application.International(xlDecimalSeparator))
    End If
    VS1 = Trim(VS1)
    If IsNumeric(VS1) Then
    r1 = CDbl(VS1)
    If ER = 0 Then
    TbWEP.ListColumns("r").DataBodyRange(1, 1).value = r1 / 100
    End If
    Else
    Me.S1.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.S2.value = "" Then
    Me.S2.BorderStyle = fmBorderStyleSingle
    Me.S2.BorderColor = vbRed
    ER = 1
    Else
    VS2 = Replace(Me.S2.value, localDec, altDec)
    If vbaSep = "." Then
    VS2 = Replace(VS2, altDec, Application.International(xlDecimalSeparator))
    End If
    VS2 = Trim(VS2)
    If IsNumeric(VS2) Then
    r2 = CDbl(VS2)
    If ER = 0 Then
    TbWEP.ListColumns("r").DataBodyRange(2, 1).value = r2 / 100
    End If
    Else
    Me.S2.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.S3.value = "" Then
    Me.S3.BorderStyle = fmBorderStyleSingle
    Me.S3.BorderColor = vbRed
    ER = 1
    Else
    VS3 = Replace(Me.S3.value, localDec, altDec)
    If vbaSep = "." Then
    VS3 = Replace(VS3, altDec, Application.International(xlDecimalSeparator))
    End If
    VS3 = Trim(VS3)
    If IsNumeric(VS3) Then
    r3 = CDbl(VS3)
    If ER = 0 Then
    TbWEP.ListColumns("r").DataBodyRange(3, 1).value = r3 / 100
    End If
    Else
    Me.S3.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.S4.value = "" Then
    Me.S4.BorderStyle = fmBorderStyleSingle
    Me.S4.BorderColor = vbRed
    ER = 1
    Else
    VS4 = Replace(Me.S4.value, localDec, altDec)
    If vbaSep = "." Then
    VS4 = Replace(VS4, altDec, Application.International(xlDecimalSeparator))
    End If
    VS4 = Trim(VS4)
    If IsNumeric(VS4) Then
    r4 = CDbl(VS4)
    If ER = 0 Then
    TbWEP.ListColumns("r4").DataBodyRange(1, 1).value = r4 / 100
    TbWEP.ListColumns("r4").DataBodyRange(2, 1).value = r4 / 100
    TbWEP.ListColumns("r4").DataBodyRange(3, 1).value = r4 / 100
    End If
    Else
    Me.S4.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.RIF.value = "" Then
    Me.RIF.BorderStyle = fmBorderStyleSingle
    Me.RIF.BorderColor = vbRed
    ER = 1
    Else
    VRIF = Replace(Me.RIF.value, localDec, altDec)
    If vbaSep = "." Then
    VRIF = Replace(VRIF, altDec, Application.International(xlDecimalSeparator))
    End If
    VRIF = Trim(VRIF)
    If IsNumeric(VRIF) Then
    YrY = CDbl(VRIF)
    If ER = 0 Then
    TbWEP.ListColumns("YrY").DataBodyRange(1, 1).value = YrY / 100
    TbWEP.ListColumns("YrY").DataBodyRange(2, 1).value = YrY / 100
    TbWEP.ListColumns("YrY").DataBodyRange(3, 1).value = YrY / 100
    End If
    Else
    Me.RIF.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.WM.value = "" Then
    Me.WM.BorderStyle = fmBorderStyleSingle
    Me.WM.BorderColor = vbRed
    ER = 1
    Else
    VWM = Replace(Me.WM.value, localDec, altDec)
    If vbaSep = "." Then
    VWM = Replace(VWM, altDec, Application.International(xlDecimalSeparator))
    End If
    VWM = Trim(VWM)
    If IsNumeric(VWM) Then
    Withdrawal = CDbl(VWM)
    If ER = 0 Then
    TbWEP.ListColumns("W").DataBodyRange(1, 1).value = Withdrawal
    TbWEP.ListColumns("W").DataBodyRange(2, 1).value = Withdrawal
    TbWEP.ListColumns("W").DataBodyRange(3, 1).value = Withdrawal
    End If
    Else
    Me.WM.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If ER = 1 Then
    Exit Sub
    End If
    
    Year = VBA.DateTime.Year(Date)
    YRT = Year + (RT - CT)
    DTEP.ListColumns("YRT").DataBodyRange(1, 1).value = YRT
    DTEP.ListColumns("YRT").DataBodyRange(2, 1).value = YRT
    DTEP.ListColumns("YRT").DataBodyRange(3, 1).value = YRT
    Me.MousePointer = fmMousePointerHourGlass
    LibR = ThisWorkbook.Sheets("Library").Range("A174").value
    Me.Caption = LibR
    
                       Set FRow = DTEP.ListColumns(Année).DataBodyRange.Find(What:=Year, LookIn:=xlValues, LookAt:=xlWhole)
                       On Error Resume Next
                       If Not FRow Is Nothing Then
                       SRow = FRow.Row
                       Else
                       LibR = ThisWorkbook.Sheets("Library").Range("A175").value
                       MsgBox (LibR), vbOK
                       Exit Sub
                       End If
                       
    DTEP.ListColumns("Stratégie1").DataBodyRange.ClearContents
    DTEP.ListColumns("Stratégie2").DataBodyRange.ClearContents
    DTEP.ListColumns("Stratégie3").DataBodyRange.ClearContents
                       
    r1 = (r1 / 100)
    r2 = (r2 / 100)
    r3 = (r3 / 100)
    r4 = (r4 / 100)
    YrY = (YrY / 100)
    n = 1
    CheckYear = 1
    
    
    Dim MonthlyGrowthAccum As Double, MonthlyGrowthDist As Double, MonthlyInflation As Double
    r1 = (1 + r1) ^ (1 / 12) - 1
    r2 = (1 + r2) ^ (1 / 12) - 1
    r3 = (1 + r3) ^ (1 / 12) - 1
    r4 = (1 + r4) ^ (1 / 12) - 1
    rinf = (1 + YrY) ^ (1 / 12) - 1
    r11 = r1 - rinf
    r21 = r2 - rinf
    r31 = r3 - rinf
    r41 = r4 - rinf
    
    YrY = (1 + YrY)

    ' Convert ages to months
    Dim MonthsAccumulation As Long, MonthsDistribution As Long
    MonthsAccumulation = (RT - CT) * 12
    MonthsDistribution = (FT - RT) * 12

    ' Phase 1: Accumulation Phase
    Dim P As Double, MonthlyContribution As Double
    FIRE_Number1 = m
    FIRE_Number2 = m
    FIRE_Number3 = m
    
    DTEP.ListColumns("Stratégie1").DataBodyRange(SRow - 1, 1).value = FIRE_Number1
    DTEP.ListColumns("Stratégie2").DataBodyRange(SRow - 1, 1).value = FIRE_Number2
    DTEP.ListColumns("Stratégie3").DataBodyRange(SRow - 1, 1).value = FIRE_Number3

    ' Start with the current accumulated amount
    Withdrawal = Withdrawal / 12
     
    For i = 1 To MonthsAccumulation
        
        FIRE_Number1 = (FIRE_Number1 + EPM) * (1 + r11) ' Monthly growth with contributions
        FIRE_Number2 = (FIRE_Number2 + EPM) * (1 + r21) ' Monthly growth with contributions
        FIRE_Number3 = (FIRE_Number3 + EPM) * (1 + r31) ' Monthly growth with contributions
        If CheckYear = 12 Then
           n = n + 1
           CheckYear = 0
           DTEP.ListColumns("Stratégie1").DataBodyRange(SRow, 1).value = FIRE_Number1
           DTEP.ListColumns("Stratégie2").DataBodyRange(SRow, 1).value = FIRE_Number2
           DTEP.ListColumns("Stratégie3").DataBodyRange(SRow, 1).value = FIRE_Number3
           SRow = SRow + 1
        End If
        CheckYear = CheckYear + 1
        
        
        
    Next i
    DTEP.ListColumns("FI1").DataBodyRange(2, 1).value = FIRE_Number1
    DTEP.ListColumns("FI2").DataBodyRange(2, 1).value = FIRE_Number2
    DTEP.ListColumns("FI3").DataBodyRange(2, 1).value = FIRE_Number3
    DTEP.ListColumns("FI1").DataBodyRange(3, 1).value = FIRE_Number1 * 1.1
    DTEP.ListColumns("FI2").DataBodyRange(3, 1).value = FIRE_Number2 * 1.1
    DTEP.ListColumns("FI3").DataBodyRange(3, 1).value = FIRE_Number3 * 1.1
    ' Phase 2: Find Required FIRE Number at Retirement

    ' Start with an estimate equal to accumulated capital
    P1 = FIRE_Number1
    P2 = FIRE_Number2
    P3 = FIRE_Number3

    ' Iterative approach: Adjust FIRENumber until balance reaches ~0 at EndAge
  
        MissingYears1 = 0
        MissingYears2 = 0
        MissingYears3 = 0
        CheckYear = 1
        ' Simulate withdrawals until EndAge
        For i = 1 To MonthsDistribution

            If P1 >= 0 Then
            P1 = (P1 - Withdrawal) * (1 + r41)
            End If
            If P2 >= 0 Then
            P2 = (P2 - Withdrawal) * (1 + r41)
            End If
            If P3 >= 0 Then
            P3 = (P3 - Withdrawal) * (1 + r41)
            End If
            
        If CheckYear = 12 Then
           n = n + 1
           CheckYear = 0
           If P1 > 0 Then
           DTEP.ListColumns("Stratégie1").DataBodyRange(SRow, 1).value = P1
           Else
           DTEP.ListColumns("Stratégie1").DataBodyRange(SRow, 1).value = ""
           End If
           If P2 > 0 Then
           DTEP.ListColumns("Stratégie2").DataBodyRange(SRow, 1).value = P2
           Else
           DTEP.ListColumns("Stratégie2").DataBodyRange(SRow, 1).value = ""
           End If
           If P3 > 0 Then
           DTEP.ListColumns("Stratégie3").DataBodyRange(SRow, 1).value = P3
           Else
           DTEP.ListColumns("Stratégie3").DataBodyRange(SRow, 1).value = ""
           End If
           SRow = SRow + 1
           
           If P1 < 0 Then
              MissingYears1 = MissingYears1 + 1
           End If
           If P2 < 0 Then
              MissingYears2 = MissingYears2 + 1
           End If
           If P3 < 0 Then
              MissingYears3 = MissingYears3 + 1
           End If
           
        End If
        CheckYear = CheckYear + 1
        Next i
    If P1 < 0 Then
       P1 = 0
    End If
    If P2 < 0 Then
       P2 = 0
    End If
    If P3 < 0 Then
       P3 = 0
    End If

    TbWEP.ListColumns("FIRE").DataBodyRange(1, 1).value = FIRE_Number1
    TbWEP.ListColumns("FIRE").DataBodyRange(2, 1).value = FIRE_Number2
    TbWEP.ListColumns("FIRE").DataBodyRange(3, 1).value = FIRE_Number3
    TbWEP.ListColumns("AR").DataBodyRange(1, 1).value = MissingYears1
    TbWEP.ListColumns("AR").DataBodyRange(2, 1).value = MissingYears2
    TbWEP.ListColumns("AR").DataBodyRange(3, 1).value = MissingYears3
    TbWEP.ListColumns("Capital").DataBodyRange(1, 1).value = P1
    TbWEP.ListColumns("Capital").DataBodyRange(2, 1).value = P2
    TbWEP.ListColumns("Capital").DataBodyRange(3, 1).value = P3
    TbWEP.ListColumns("FIRE").DataBodyRange.NumberFormat = "# ### ### ### ##0" & CRY
    TbWEP.ListColumns("Capital").DataBodyRange.NumberFormat = "# ### ### ### ##0" & CRY


series11.Format.Line.Visible = msoTrue
series11.MarkerStyle = xlMarkerStyleNone ' Remove markers
series11.Format.Line.Weight = 1 ' Set line width to 1.25
series11.Smooth = True ' Enable smoothed line
series11.Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
series1.Format.Line.Visible = msoTrue
series1.Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
series1.MarkerStyle = xlMarkerStyleNone
series22.Format.Line.Visible = msoTrue
series22.MarkerStyle = xlMarkerStyleNone ' Remove markers
series22.Format.Line.Weight = 1 ' Set line width to 1.25
series22.Smooth = True ' Enable smoothed line
series22.Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
series2.Format.Line.Visible = msoTrue
series2.Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
series2.MarkerStyle = xlMarkerStyleNone
series33.Format.Line.Visible = msoTrue
series33.MarkerStyle = xlMarkerStyleNone ' Remove markers
series33.Format.Line.Weight = 1 ' Set line width to 1.25
series33.Smooth = True ' Enable smoothed line
series33.Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
series3.Format.Line.Visible = msoTrue
series3.Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
series3.MarkerStyle = xlMarkerStyleNone
    
    
    
    
    
    series1.Points(2).ApplyDataLabels
    Set dataLabel = series1.Points(2).dataLabel
    dataLabel.Position = xlLabelPositionLeft ' Set position to left of the point
    dataLabel.ShowValue = True
    dataLabel.NumberFormat = "# ### ### ### ##0" & CRY ' Format the Y value as currency
    dataLabel.ShowCategoryName = False
    dataLabel.Font.Name = "Aptos Narrow" ' Set font to Aptos Narrow
    dataLabel.Font.Size = 12 ' Set font size to 12
    dataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(197, 78, 248)
    
    series2.Points(2).ApplyDataLabels
    Set dataLabel = series2.Points(2).dataLabel
    dataLabel.Position = xlLabelPositionLeft ' Set position to left of the point
    dataLabel.ShowValue = True
    dataLabel.NumberFormat = "# ### ### ### ##0" & CRY ' Format the Y value as currency
    dataLabel.ShowCategoryName = False
    dataLabel.Font.Name = "Aptos Narrow" ' Set font to Aptos Narrow
    dataLabel.Font.Size = 12 ' Set font size to 12
    dataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 0)
    
    series3.Points(2).ApplyDataLabels
    Set dataLabel = series3.Points(2).dataLabel
    dataLabel.Position = xlLabelPositionLeft ' Set position to left of the point
    dataLabel.ShowValue = True
    dataLabel.NumberFormat = "# ### ### ### ##0" & CRY ' Format the Y value as currency
    dataLabel.ShowCategoryName = False
    dataLabel.Font.Name = "Aptos Narrow" ' Set font to Aptos Narrow
    dataLabel.Font.Size = 12 ' Set font size to 12
    dataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(250, 97, 72)
    TbWEP.ListColumns("CHK").DataBodyRange(1, 1).value = 1
    TbWEP.ListColumns("CHK").DataBodyRange(2, 1).value = 1
    TbWEP.ListColumns("CHK").DataBodyRange(3, 1).value = 1
    TbWEP.ListColumns("Curve").DataBodyRange(1, 1).value = 1
    TbWEP.ListColumns("Curve").DataBodyRange(2, 1).value = 1
    TbWEP.ListColumns("Curve").DataBodyRange(3, 1).value = 1
    FIRE_Number3 = TbWEP.ListColumns("FIRE").DataBodyRange(3, 1).value * 1.2
    Set yAxis = chartObj.Chart.Axes(xlValue)
    yAxis.MaximumScale = FIRE_Number3
    yAxis.MinimumScale = 0
    Set xAxis = chartObj.Chart.Axes(xlCategory)
    xAxis.MaximumScale = (FT + 1 + 2025) - CT
    xAxis.MinimumScale = 2025
    Set SHStrat = ThisWorkbook.Sheets("DashBoard").Shapes("ShapeStrat11")
    SHStrat.Line.ForeColor.RGB = RGB(170, 220, 215) ' Red border
    SHStrat.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(170, 220, 215)
    Set SHStrat = ThisWorkbook.Sheets("DashBoard").Shapes("ShapeStrat22")
    SHStrat.Line.ForeColor.RGB = RGB(170, 220, 215) ' Red border
    SHStrat.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(170, 220, 215)
    Set SHStrat = ThisWorkbook.Sheets("DashBoard").Shapes("ShapeStrat33")
    SHStrat.Line.ForeColor.RGB = RGB(170, 220, 215) ' Red border
    SHStrat.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(170, 220, 215)
    Me.MousePointer = fmMousePointerDefault
    LibR = ThisWorkbook.Sheets("Library").Range("A144").value
    Me.Caption = LibR
Unload Me

End Sub
Private Sub Valider_Click()
If Me.CheckInf = True Then
   Validerinf
ElseIf Me.CheckInf = False Then
   ValiderNoInf
End If
End Sub
Private Sub Validerinf_Click()
Dim LibR As String
ER = 0
    Me.Strat2.BorderStyle = fmBorderStyleNone
    Me.Strat2.BorderColor = &HC0C0C0
    Me.LS2.BorderStyle = fmBorderStyleNone
    Me.LS2.BorderColor = &HC0C0C0
    
    If Me.Tf.value = "" Then
        Me.Tf.BorderStyle = fmBorderStyleSingle
        Me.Tf.BorderColor = vbRed
    ER = 1
    Else
    If IsNumeric(Me.Tf.value) Then
    FT = Me.Tf.value + 0
    If FT <> Int(FT) Then
    Me.Tf.ForeColor = vbRed
    ER = 1
    End If
    If ER = 0 Then
    TbWEP.ListColumns("tf").DataBodyRange(1, 1).value = FT
    TbWEP.ListColumns("tf").DataBodyRange(2, 1).value = FT
    TbWEP.ListColumns("tf").DataBodyRange(3, 1).value = FT
    End If
    Else
    Me.Tf.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    
    If Me.Tc.value = "" Then
    Me.Tc.BorderStyle = fmBorderStyleSingle
    Me.Tc.BorderColor = vbRed
    ER = 1
    Else
    If IsNumeric(Me.Tc.value) Then
    CT = Me.Tc.value + 0
    If CT <> Int(CT) Then
    Me.Tc.ForeColor = vbRed
    ER = 1
    End If
    If ER = 0 Then
    TbWEP.ListColumns("tc").DataBodyRange(1, 1).value = CT
    TbWEP.ListColumns("tc").DataBodyRange(2, 1).value = CT
    TbWEP.ListColumns("tc").DataBodyRange(3, 1).value = CT
    End If
    Else
    Me.Tc.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.Tr.value = "" Then
    Me.Tr.BorderStyle = fmBorderStyleSingle
    Me.Tr.BorderColor = vbRed
    ER = 1
    Else
    If IsNumeric(Me.Tr.value) Then
    RT = Me.Tr.value + 0
    If RT <> Int(RT) Then
    Me.Tr.ForeColor = vbRed
    ER = 1
    End If
    If ER = 0 Then
    TbWEP.ListColumns("tr").DataBodyRange(1, 1).value = RT
    TbWEP.ListColumns("tr").DataBodyRange(2, 1).value = RT
    TbWEP.ListColumns("tr").DataBodyRange(3, 1).value = RT
    End If
    Else
    Me.Tr.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.SB.value = "" Then
    Me.SB.BorderStyle = fmBorderStyleSingle
    Me.SB.BorderColor = vbRed
    ER = 1
    Else
    VSB = Replace(Me.SB.value, localDec, altDec)
    If vbaSep = "." Then
    VSB = Replace(VSB, altDec, Application.International(xlDecimalSeparator))
    End If
    VSB = Trim(VSB)
    If IsNumeric(VSB) Then
    SAB = CDbl(VSB)
    If ER = 0 Then
    TbWEP.ListColumns("SB").DataBodyRange(1, 1).value = SAB
    TbWEP.ListColumns("SB").DataBodyRange(2, 1).value = SAB
    TbWEP.ListColumns("SB").DataBodyRange(3, 1).value = SAB
    End If
    Else
    Me.SB.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.Montant.value = "" Then
    m = 0
    Else
    Vm = Replace(Me.Montant.value, localDec, altDec)
    If vbaSep = "." Then
    Vm = Replace(Vm, altDec, Application.International(xlDecimalSeparator))
    End If
    Vm = Trim(Vm)
    If IsNumeric(Vm) Then
    m = CDbl(Vm)
    If ER = 0 Then
    TbWEP.ListColumns("M").DataBodyRange(1, 1).value = m
    TbWEP.ListColumns("M").DataBodyRange(2, 1).value = m
    TbWEP.ListColumns("M").DataBodyRange(3, 1).value = m
    End If
    Else
    Me.Montant.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.MEP.value = "" Then
    EPM = 0
    Else
    VMEPM = Replace(Me.MEP.value, localDec, altDec)
    If vbaSep = "." Then
    VMEPM = Replace(VMEPM, altDec, Application.International(xlDecimalSeparator))
    End If
    VMEPM = Trim(VMEPM)
    If IsNumeric(VMEPM) Then
    EPM = CDbl(VEPM)
    If ER = 0 Then
    TbWEP.ListColumns("MEP").DataBodyRange(1, 1).value = EPM
    TbWEP.ListColumns("MEP").DataBodyRange(2, 1).value = EPM
    TbWEP.ListColumns("MEP").DataBodyRange(3, 1).value = EPM
    End If
    Else
    Me.MEP.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If m = 0 And EPM = 0 Then
    Me.Montant.BorderStyle = fmBorderStyleSingle
    Me.Montant.BorderColor = vbRed
    Me.MEP.BorderStyle = fmBorderStyleSingle
    Me.MEP.BorderColor = vbRed
    End If
    
    
    If Me.S1.value = "" Then
    Me.S1.BorderStyle = fmBorderStyleSingle
    Me.S1.BorderColor = vbRed
    ER = 1
    Else
    VS1 = Replace(Me.S1.value, localDec, altDec)
    If vbaSep = "." Then
    VS1 = Replace(VS1, altDec, Application.International(xlDecimalSeparator))
    End If
    VS1 = Trim(VS1)
    If IsNumeric(VS1) Then
    r1 = CDbl(VS1)
    If ER = 0 Then
    TbWEP.ListColumns("r").DataBodyRange(1, 1).value = r1 / 100
    End If
    Else
    Me.S1.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.S2.value = "" Then
    Me.S2.BorderStyle = fmBorderStyleSingle
    Me.S2.BorderColor = vbRed
    ER = 1
    Else
    VS2 = Replace(Me.S2.value, localDec, altDec)
    If vbaSep = "." Then
    VS2 = Replace(VS2, altDec, Application.International(xlDecimalSeparator))
    End If
    VS2 = Trim(VS2)
    If IsNumeric(VS2) Then
    r2 = CDbl(VS2)
    If ER = 0 Then
    TbWEP.ListColumns("r").DataBodyRange(2, 1).value = r2 / 100
    End If
    Else
    Me.S2.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.S3.value = "" Then
    Me.S3.BorderStyle = fmBorderStyleSingle
    Me.S3.BorderColor = vbRed
    ER = 1
    Else
    VS3 = Replace(Me.S3.value, localDec, altDec)
    If vbaSep = "." Then
    VS3 = Replace(VS3, altDec, Application.International(xlDecimalSeparator))
    End If
    VS3 = Trim(VS3)
    If IsNumeric(VS3) Then
    r3 = CDbl(VS3)
    If ER = 0 Then
    TbWEP.ListColumns("r").DataBodyRange(3, 1).value = r3 / 100
    End If
    Else
    Me.S3.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.S4.value = "" Then
    Me.S4.BorderStyle = fmBorderStyleSingle
    Me.S4.BorderColor = vbRed
    ER = 1
    Else
    VS4 = Replace(Me.S4.value, localDec, altDec)
    If vbaSep = "." Then
    VS4 = Replace(VS4, altDec, Application.International(xlDecimalSeparator))
    End If
    VS4 = Trim(VS4)
    If IsNumeric(VS4) Then
    r4 = CDbl(VS4)
    If ER = 0 Then
    TbWEP.ListColumns("r4").DataBodyRange(1, 1).value = r4 / 100
    TbWEP.ListColumns("r4").DataBodyRange(2, 1).value = r4 / 100
    TbWEP.ListColumns("r4").DataBodyRange(3, 1).value = r4 / 100
    End If
    Else
    Me.S4.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.RIF.value = "" Then
    Me.RIF.BorderStyle = fmBorderStyleSingle
    Me.RIF.BorderColor = vbRed
    ER = 1
    Else
    VRIF = Replace(Me.RIF.value, localDec, altDec)
    If vbaSep = "." Then
    VRIF = Replace(VRIF, altDec, Application.International(xlDecimalSeparator))
    End If
    VRIF = Trim(VRIF)
    If IsNumeric(VRIF) Then
    YrY = CDbl(VRIF)
    If ER = 0 Then
    TbWEP.ListColumns("YrY").DataBodyRange(1, 1).value = YrY / 100
    TbWEP.ListColumns("YrY").DataBodyRange(2, 1).value = YrY / 100
    TbWEP.ListColumns("YrY").DataBodyRange(3, 1).value = YrY / 100
    End If
    Else
    Me.RIF.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If Me.WM.value = "" Then
    Me.WM.BorderStyle = fmBorderStyleSingle
    Me.WM.BorderColor = vbRed
    ER = 1
    Else
    VWM = Replace(Me.WM.value, localDec, altDec)
    If vbaSep = "." Then
    VWM = Replace(VWM, altDec, Application.International(xlDecimalSeparator))
    End If
    VWM = Trim(VWM)
    If IsNumeric(VWM) Then
    Withdrawal = CDbl(VWM)
    If ER = 0 Then
    TbWEP.ListColumns("W").DataBodyRange(1, 1).value = Withdrawal
    TbWEP.ListColumns("W").DataBodyRange(2, 1).value = Withdrawal
    TbWEP.ListColumns("W").DataBodyRange(3, 1).value = Withdrawal
    End If
    Else
    Me.WM.ForeColor = vbRed
    ER = 1
    End If
    End If
    
    If ER = 1 Then
    Exit Sub
    End If
    
    Year = VBA.DateTime.Year(Date)
    YRT = Year + (RT - CT)
    DTEP.ListColumns("YRT").DataBodyRange(1, 1).value = YRT
    DTEP.ListColumns("YRT").DataBodyRange(2, 1).value = YRT
    DTEP.ListColumns("YRT").DataBodyRange(3, 1).value = YRT
    Me.MousePointer = fmMousePointerHourGlass
    LibR = ThisWorkbook.Sheets("Library").Range("A174").value
    Me.Caption = LibR
    
                       Set FRow = DTEP.ListColumns(Année).DataBodyRange.Find(What:=Year, LookIn:=xlValues, LookAt:=xlWhole)
                       On Error Resume Next
                       If Not FRow Is Nothing Then
                       SRow = FRow.Row
                       Else
                       LibR = ThisWorkbook.Sheets("Library").Range("A175").value
                       MsgBox (LibR), vbOK
                       Exit Sub
                       End If
                       
    DTEP.ListColumns("Stratégie1").DataBodyRange.ClearContents
    DTEP.ListColumns("Stratégie2").DataBodyRange.ClearContents
    DTEP.ListColumns("Stratégie3").DataBodyRange.ClearContents
                       
    r1 = (r1 / 100)
    r2 = (r2 / 100)
    r3 = (r3 / 100)
    r4 = (r4 / 100)
    YrY = (YrY / 100)
    n = 1
    CheckYear = 1
    
    
    Dim MonthlyGrowthAccum As Double, MonthlyGrowthDist As Double, MonthlyInflation As Double
    r1 = (1 + r1) ^ (1 / 12) - 1
    r2 = (1 + r2) ^ (1 / 12) - 1
    r3 = (1 + r3) ^ (1 / 12) - 1
    r4 = (1 + r4) ^ (1 / 12) - 1
    YrY = (1 + YrY)

    ' Convert ages to months
    Dim MonthsAccumulation As Long, MonthsDistribution As Long
    MonthsAccumulation = (RT - CT) * 12
    MonthsDistribution = (FT - RT) * 12

    ' Phase 1: Accumulation Phase
    Dim P As Double, MonthlyContribution As Double
    FIRE_Number1 = m
    FIRE_Number2 = m
    FIRE_Number3 = m
    
    DTEP.ListColumns("Stratégie1").DataBodyRange(SRow - 1, 1).value = FIRE_Number1
    DTEP.ListColumns("Stratégie2").DataBodyRange(SRow - 1, 1).value = FIRE_Number2
    DTEP.ListColumns("Stratégie3").DataBodyRange(SRow - 1, 1).value = FIRE_Number3

    ' Start with the current accumulated amount
    Withdrawal = Withdrawal / 12
     
    For i = 1 To MonthsAccumulation
        
        FIRE_Number1 = (FIRE_Number1 + EPM) * (1 + r1) ' Monthly growth with contributions
        FIRE_Number2 = (FIRE_Number2 + EPM) * (1 + r2) ' Monthly growth with contributions
        FIRE_Number3 = (FIRE_Number3 + EPM) * (1 + r3) ' Monthly growth with contributions
        If CheckYear = 12 Then
           n = n + 1
           EPM = EPM * YrY
           Withdrawal = Withdrawal * YrY
           CheckYear = 0
           DTEP.ListColumns("Stratégie1").DataBodyRange(SRow, 1).value = FIRE_Number1
           DTEP.ListColumns("Stratégie2").DataBodyRange(SRow, 1).value = FIRE_Number2
           DTEP.ListColumns("Stratégie3").DataBodyRange(SRow, 1).value = FIRE_Number3
           SRow = SRow + 1
        End If
        CheckYear = CheckYear + 1
        
        
        
    Next i
    DTEP.ListColumns("FI1").DataBodyRange(2, 1).value = FIRE_Number1
    DTEP.ListColumns("FI2").DataBodyRange(2, 1).value = FIRE_Number2
    DTEP.ListColumns("FI3").DataBodyRange(2, 1).value = FIRE_Number3
    DTEP.ListColumns("FI1").DataBodyRange(3, 1).value = FIRE_Number1 * 1.1
    DTEP.ListColumns("FI2").DataBodyRange(3, 1).value = FIRE_Number2 * 1.1
    DTEP.ListColumns("FI3").DataBodyRange(3, 1).value = FIRE_Number3 * 1.1
    ' Phase 2: Find Required FIRE Number at Retirement

    ' Start with an estimate equal to accumulated capital
    P1 = FIRE_Number1
    P2 = FIRE_Number2
    P3 = FIRE_Number3

    ' Iterative approach: Adjust FIRENumber until balance reaches ~0 at EndAge
  
        MissingYears1 = 0
        MissingYears2 = 0
        MissingYears3 = 0
        CheckYear = 1
        ' Simulate withdrawals until EndAge
        For i = 1 To MonthsDistribution

            If P1 >= 0 Then
            P1 = (P1 - Withdrawal) * (1 + r4)
            End If
            If P2 >= 0 Then
            P2 = (P2 - Withdrawal) * (1 + r4)
            End If
            If P3 >= 0 Then
            P3 = (P3 - Withdrawal) * (1 + r4)
            End If
            
        If CheckYear = 12 Then
           n = n + 1
           Withdrawal = Withdrawal * YrY
           CheckYear = 0
           If P1 > 0 Then
           DTEP.ListColumns("Stratégie1").DataBodyRange(SRow, 1).value = P1
           Else
           DTEP.ListColumns("Stratégie1").DataBodyRange(SRow, 1).value = ""
           End If
           If P2 > 0 Then
           DTEP.ListColumns("Stratégie2").DataBodyRange(SRow, 1).value = P2
           Else
           DTEP.ListColumns("Stratégie2").DataBodyRange(SRow, 1).value = ""
           End If
           If P3 > 0 Then
           DTEP.ListColumns("Stratégie3").DataBodyRange(SRow, 1).value = P3
           Else
           DTEP.ListColumns("Stratégie3").DataBodyRange(SRow, 1).value = ""
           End If
           SRow = SRow + 1
           
           If P1 < 0 Then
              MissingYears1 = MissingYears1 + 1
           End If
           If P2 < 0 Then
              MissingYears2 = MissingYears2 + 1
           End If
           If P3 < 0 Then
              MissingYears3 = MissingYears3 + 1
           End If
           
        End If
        CheckYear = CheckYear + 1
        Next i
    If P1 < 0 Then
       P1 = 0
    End If
    If P2 < 0 Then
       P2 = 0
    End If
    If P3 < 0 Then
       P3 = 0
    End If

    TbWEP.ListColumns("FIRE").DataBodyRange(1, 1).value = FIRE_Number1
    TbWEP.ListColumns("FIRE").DataBodyRange(2, 1).value = FIRE_Number2
    TbWEP.ListColumns("FIRE").DataBodyRange(3, 1).value = FIRE_Number3
    TbWEP.ListColumns("AR").DataBodyRange(1, 1).value = MissingYears1
    TbWEP.ListColumns("AR").DataBodyRange(2, 1).value = MissingYears2
    TbWEP.ListColumns("AR").DataBodyRange(3, 1).value = MissingYears3
    TbWEP.ListColumns("Capital").DataBodyRange(1, 1).value = P1
    TbWEP.ListColumns("Capital").DataBodyRange(2, 1).value = P2
    TbWEP.ListColumns("Capital").DataBodyRange(3, 1).value = P3
    TbWEP.ListColumns("FIRE").DataBodyRange.NumberFormat = "# ### ### ### ##0" & CRY
    TbWEP.ListColumns("Capital").DataBodyRange.NumberFormat = "# ### ### ### ##0" & CRY


series11.Format.Line.Visible = msoTrue
series11.MarkerStyle = xlMarkerStyleNone ' Remove markers
series11.Format.Line.Weight = 1 ' Set line width to 1.25
series11.Smooth = True ' Enable smoothed line
series11.Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
series1.Format.Line.Visible = msoTrue
series1.Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
series1.MarkerStyle = xlMarkerStyleNone
series22.Format.Line.Visible = msoTrue
series22.MarkerStyle = xlMarkerStyleNone ' Remove markers
series22.Format.Line.Weight = 1 ' Set line width to 1.25
series22.Smooth = True ' Enable smoothed line
series22.Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
series2.Format.Line.Visible = msoTrue
series2.Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
series2.MarkerStyle = xlMarkerStyleNone
series33.Format.Line.Visible = msoTrue
series33.MarkerStyle = xlMarkerStyleNone ' Remove markers
series33.Format.Line.Weight = 1 ' Set line width to 1.25
series33.Smooth = True ' Enable smoothed line
series33.Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
series3.Format.Line.Visible = msoTrue
series3.Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
series3.MarkerStyle = xlMarkerStyleNone
    
    
    
    
    
    series1.Points(2).ApplyDataLabels
    Set dataLabel = series1.Points(2).dataLabel
    dataLabel.Position = xlLabelPositionLeft ' Set position to left of the point
    dataLabel.ShowValue = True
    dataLabel.NumberFormat = "# ### ### ### ##0" & CRY ' Format the Y value as currency
    dataLabel.ShowCategoryName = False
    dataLabel.Font.Name = "Aptos Narrow" ' Set font to Aptos Narrow
    dataLabel.Font.Size = 12 ' Set font size to 12
    dataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(197, 78, 248)
    
    series2.Points(2).ApplyDataLabels
    Set dataLabel = series2.Points(2).dataLabel
    dataLabel.Position = xlLabelPositionLeft ' Set position to left of the point
    dataLabel.ShowValue = True
    dataLabel.NumberFormat = "# ### ### ### ##0" & CRY ' Format the Y value as currency
    dataLabel.ShowCategoryName = False
    dataLabel.Font.Name = "Aptos Narrow" ' Set font to Aptos Narrow
    dataLabel.Font.Size = 12 ' Set font size to 12
    dataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 0)
    
    series3.Points(2).ApplyDataLabels
    Set dataLabel = series3.Points(2).dataLabel
    dataLabel.Position = xlLabelPositionLeft ' Set position to left of the point
    dataLabel.ShowValue = True
    dataLabel.NumberFormat = "# ### ### ### ##0" & CRY ' Format the Y value as currency
    dataLabel.ShowCategoryName = False
    dataLabel.Font.Name = "Aptos Narrow" ' Set font to Aptos Narrow
    dataLabel.Font.Size = 12 ' Set font size to 12
    dataLabel.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(250, 97, 72)
    TbWEP.ListColumns("CHK").DataBodyRange(1, 1).value = 1
    TbWEP.ListColumns("CHK").DataBodyRange(2, 1).value = 1
    TbWEP.ListColumns("CHK").DataBodyRange(3, 1).value = 1
    TbWEP.ListColumns("Curve").DataBodyRange(1, 1).value = 1
    TbWEP.ListColumns("Curve").DataBodyRange(2, 1).value = 1
    TbWEP.ListColumns("Curve").DataBodyRange(3, 1).value = 1
    FIRE_Number3 = TbWEP.ListColumns("FIRE").DataBodyRange(3, 1).value * 1.2
    Set yAxis = chartObj.Chart.Axes(xlValue)
    yAxis.MaximumScale = FIRE_Number3
    yAxis.MinimumScale = 0
    Set xAxis = chartObj.Chart.Axes(xlCategory)
    xAxis.MaximumScale = (FT + 1 + 2025) - CT
    xAxis.MinimumScale = 2025
    Set SHStrat = ThisWorkbook.Sheets("DashBoard").Shapes("ShapeStrat11")
    SHStrat.Line.ForeColor.RGB = RGB(170, 220, 215) ' Red border
    SHStrat.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(170, 220, 215)
    Set SHStrat = ThisWorkbook.Sheets("DashBoard").Shapes("ShapeStrat22")
    SHStrat.Line.ForeColor.RGB = RGB(170, 220, 215) ' Red border
    SHStrat.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(170, 220, 215)
    Set SHStrat = ThisWorkbook.Sheets("DashBoard").Shapes("ShapeStrat33")
    SHStrat.Line.ForeColor.RGB = RGB(170, 220, 215) ' Red border
    SHStrat.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(170, 220, 215)
    Me.MousePointer = fmMousePointerDefault
    LibR = ThisWorkbook.Sheets("Library").Range("A144").value
    Me.Caption = LibR
Unload Me

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

LibR = ThisWorkbook.Sheets("Library").Range("A144").value
Me.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A145").value
Me.LTc.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A146").value
Me.LTr.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A147").value
Me.LTf.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A148").value
Me.LSB.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A149").value
Me.LMontant.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A150").value
Me.LMEP.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A151").value
Me.QMEP.Caption = LibR
Me.CHKWM.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A153").value
Me.Label43.Caption = LibR
Me.Label41.Caption = LibR
Me.Label42.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A154").value
Me.Label66.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A155").value
Me.CheckInf.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A156").value
Me.Label49.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A157").value
Me.Label50.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A158").value
Me.LS1.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A159").value
Me.LS2.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A160").value
Me.LS3.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A161").value
Me.Label51.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A162").value
Me.LS4.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A163").value
Me.LRIF.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A164").value
Me.LWM.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A165").value
Me.Label52.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A166").value
Me.Label59.Caption = LibR
Me.Label65.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A167").value
Me.Strat2.Caption = LibR
Me.Label61.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A168").value
Me.Label54.Caption = LibR
Me.Label60.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A169").value
Me.Label53.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A170").value
Me.LAA1.Caption = LibR
Me.LAA2.Caption = LibR
Me.LAA3.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A171").value
Me.Quitt.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A172").value
Me.CLCFIRE.Caption = LibR
LibR = ThisWorkbook.Sheets("Library").Range("A173").value
Me.Valider.Caption = LibR


Set tbCCY = ThisWorkbook.Sheets("WB_Data").ListObjects("tbCCY")
Set chartObj = ThisWorkbook.Sheets("DashBoard").ChartObjects("Chart Strategique")
Set series1 = chartObj.Chart.SeriesCollection("Stratégie 1")
Set series2 = chartObj.Chart.SeriesCollection("Stratégie 2")
Set series3 = chartObj.Chart.SeriesCollection("Stratégie 3")
Set series11 = chartObj.Chart.SeriesCollection("Stratégie1")
Set series22 = chartObj.Chart.SeriesCollection("Stratégie2")
Set series33 = chartObj.Chart.SeriesCollection("Stratégie3")
Set TbWEP = ThisWorkbook.Sheets("WB_Data").ListObjects("TbDataEP")
Set DTEP = ThisWorkbook.Sheets("ControlChart").ListObjects("DTEP")
Me.Tf.value = TbWEP.ListColumns("tf").DataBodyRange(1, 1).value
Me.Tc.value = TbWEP.ListColumns("tc").DataBodyRange(1, 1).value
Me.Tr.value = TbWEP.ListColumns("tr").DataBodyRange(1, 1).value
Me.SB.value = TbWEP.ListColumns("SB").DataBodyRange(1, 1).value
Me.Montant.value = TbWEP.ListColumns("M").DataBodyRange(1, 1).value
Me.MEP.value = TbWEP.ListColumns("MEP").DataBodyRange(1, 1).value
Me.S1.value = TbWEP.ListColumns("r").DataBodyRange(1, 1).value * 100
Me.S2.value = TbWEP.ListColumns("r").DataBodyRange(2, 1).value * 100
Me.S3.value = TbWEP.ListColumns("r").DataBodyRange(3, 1).value * 100
Set tbh = ThisWorkbook.Sheets("WB_Data").ListObjects("TbHead")
Année = tbh.ListColumns("Header1").DataBodyRange(1, 1).value
Me.S4.value = TbWEP.ListColumns("r4").DataBodyRange(1, 1).value * 100
Me.RIF.value = TbWEP.ListColumns("YrY").DataBodyRange(1, 1).value * 100
Me.WM.value = TbWEP.ListColumns("W").DataBodyRange(1, 1).value
CRY = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
Me.SIG1.Caption = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
Me.SIG2.Caption = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
Me.SIG3.Caption = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
Me.SIG4.Caption = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
Me.LFDD1.Caption = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
Me.LFDD2.Caption = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
Me.LFDD3.Caption = tbCCY.ListColumns("Current").DataBodyRange(1, 1).value
Me.LFDD1.ForeColor = vbBlack
Me.LFDD2.ForeColor = vbBlack
Me.LFDD3.ForeColor = vbBlack
Me.LF1.Caption = ""
Me.LF2.Caption = ""
Me.LF3.Caption = ""
Me.LA1.Caption = ""
Me.LA2.Caption = ""
Me.LA3.Caption = ""
Me.LAA1.ForeColor = vbBlack
Me.LAA2.ForeColor = vbBlack
Me.LAA3.ForeColor = vbBlack
A = 0
Q = 0
End Sub

Private Sub WM_Change()
Me.WM.BorderStyle = fmBorderStyleNone
Me.WM.ForeColor = vbWhite
End Sub
