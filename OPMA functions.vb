' Open Psychometric Meta-Analyis (Excel)
' Created by Brenton M. Wiernik
' version 1.0.1

'    Open Psychometric Meta-Analysis (Excel) -- VBA scripts for conducting psychometric
'    meta-analysis using Microsoft Excel.
'    Copyright (C) 2017 Brenton M. Wiernik.

'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.

'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.

'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.


' Variable list:
' k = number of effect sizes
' nRyy = number of reliability (ryy) values
' R = matrix of effect sizes and sample sizes
' RY = matrix of ryy values and frequencies

' Ntotal = total sample size (N)
' sumR = weighted sumR of uncorrected d values
' meanR = weighted average uncorrected d value

' SampErrVar = Expected sampling error
' ObsSSQ = Observed Sums of Squares
' ObsVar = Observed variance of effect sizes
' SDobs = Observed SD of effect sizes
' PercentVarSamp = Percent of observed effect size variance due to sampling error

' SumRyy = Sum of y measure reliabilities
' SumRyyFreq = Sum  of reliability frequencies
' SumRyySq = Sum of squared y measure reliabilities
' MeanRyy = Mean y measure reliability
' SDRyy = SD of y measure reliabilities
' SumQualy = Sum of y measure qualities (sqrt of reliability)
' MeanQualy = Mean y measure quality (sqrt of reliability)
' SDQualy = SD of y measure qualities (sqrt of reliabilities)

' Rho = estimated mean true d effect size

' Ratten = Attenuated d value for a particular ryy value
' RattenWeighted = Weighted attenuated d value for a particular ryy value
' SumRatten = Sum of attenuated d values
' SumRattenSq = Sum of squared attenuated d values
' ArtVar = Expected variance due to artifacts

' ResVar = Residual Variance of effect sizes
' SDres = Residual SD of effect sizes
' SDpred = Predicted SD of effect sizes
' PerVarAcc = Percent of variance in effect sizes accounted for

' SDrho = True effect standard deviation

' SEmeanR = Standard error of mean d
' df = Degrees of freedom for confidence interval when k < 30
' Crit = Critical value for confidence interval when k < 30
' UpCIdelta = Upper value of 80% confidence interval for delta
' LoCIdelta = Lower value of 80% confidence interval for delta
' UpCImeanR = Upper value of 80% confidence interval for mean d
' LoCImeanD = Lower value of 80% confidence interval for mean d

' UpCVdelta = Upper value of 80% credibility interval for delta
' LoCVdelta = Lower value of 80% credibility interval for delta
' UpCVmeanD = Upper value of 80% credibility interval for mean d
' LoCVmeanD = Lower value of 80% credibility interval for mean d

Function IsMac() As Boolean
#If Mac Then
    IsMac = True
#End If
End Function

Function Is64BitOffice() As Boolean
#If Win64 Then
    Is64BitOffice = True
#End If
End Function

Function Excelversion() As Double
'Win Excel versions are always a whole number (15)
'Mac Excel versions show also the number of the update (15.29)
    Excelversion = Val(Application.Version)
End Function

Function NeedCompatMode() As Boolean
  #If IsMac Then
    #If Excelversion < 15 Then
       NeedCompatMode = True
    #Else
       NeedCompatMode = False
    #End If
  #Else
    #If Excelversion < 14 Then
       NeedCompatMode = True
    #Else
       NeedCompatMode = False
    #End If
  #End If
End Function

Function TruncNorm(Min As Double, Max As Double, Mean As Variant, SD As Variant)
  Randomize ' Initialize the random number generator
  CompatMode = NeedCompatMode
    
  If CompatMode Then
    MinD = WorksheetFunction.NormDist(Min, Mean, SD, True)
    MaxD = Rnd() * (WorksheetFunction.NormDist(Max, Mean, SD, True) - WorksheetFunction.NormDist(Min, Mean, SD, True))
    RangeD = MinD + MaxD
    TruncNorm = WorksheetFunction.NormInv(RangeD, Mean, SD)
  Else
    MinD = WorksheetFunction.Norm_Dist(Min, Mean, SD, True)
    MaxD = Rnd() * (WorksheetFunction.Norm_Dist(Max, Mean, SD, True) - WorksheetFunction.Norm_Dist(Min, Mean, SD, True))
    RangeD = MinD + MaxD
    TruncNorm = WorksheetFunction.Norm_Inv(RangeD, Mean, SD)
  End If
End Function

Sub MetaAnalysisR()

' Turn off screen updating for efficiency
Application.ScreenUpdating = False

' Set up control parameters based on OS version and options set via worksheet controls
CompatMode = NeedCompatMode

Dim alert As Integer

' Set alert flag counter to zero and reset all flags
flags = 0
FlagRxxOver = vbNullString
FlagQualxUnder = vbNullString
FlagNoSDQualx = vbNullString
FlagQualyUnder = vbNullString
FlagNoSDQualy = vbNullString

' ==================================
' === Choose meta-analytic model ===
' ==================================

' TODO: Add alternative weighting methods
' === Weighting method options =====
' If Worksheets("Correlations").Shapes("WtTotal").ControlFormat.Value = 1 Then
'     Weights = "Total"
' ElseIf Worksheets("Correlations").Shapes("WtUnit").ControlFormat.Value = 1 Then
'     Weights = "Unit"
' ElseIf Worksheets("Correlations").Shapes("WtInvSamp").ControlFormat.Value = 1 Then
'    Weights = "InvSamp"
' Else
'     alert = MsgBox ("Error: Please select a weighting method", vbCritical, "Choose an option")
'     Exit Sub
' End If

' === Artifact correction options ===

' ===== Reliability in X (Rxx) ======
If Worksheets("rxx").Shapes("NoDistx").ControlFormat.Value = 1 Then
  CorrectRxx = False
ElseIf Worksheets("rxx").Shapes("NewDistx").ControlFormat.Value = 1 Then
  CorrectRxx = True
  SpecDistx = False
ElseIf Worksheets("rxx").Shapes("SpecDistx").ControlFormat.Value = 1 Then
  CorrectRxx = True
  SpecDistx = True
Else
  alert = MsgBox("Error: Please an option for correcting for unreliability in X", vbCritical, "Choose an option")
  Exit Sub
End If

If Worksheets("rxx").Shapes("RelUnrestx").ControlFormat.Value = 1 Then
  RelUnrestx = True
ElseIf Worksheets("rxx").Shapes("RelRestx").ControlFormat.Value = 1 Then
  RelUnrestx = False
ElseIf CorrectRxx Then
  alert = MsgBox("Error: Please indicate whether reliability values for X are from the restricted (incumbent) or unrestricted (applicant) group", vbCritical, "Choose an option")
  Exit Sub
End If

' ===== Reliability in Y (Ryy) ======
' TODO: Allow specification of whether reliabilities for Y are from restricted or unrestricted population
If Worksheets("ryy").Shapes("NoDisty").ControlFormat.Value = 1 Then
  CorrectRyy = False
ElseIf Worksheets("ryy").Shapes("NewDisty").ControlFormat.Value = 1 Then
  CorrectRyy = True
  SpecDisty = False
ElseIf Worksheets("ryy").Shapes("SpecDisty").ControlFormat.Value = 1 Then
  CorrectRyy = True
  SpecDisty = True
Else
  alert = MsgBox("Error: Please an option for correcting for unreliability in Y", vbCritical, "Choose an option")
  Exit Sub
End If

' ======== Range restriction ========
If Worksheets("RR").Shapes("NoDistu").ControlFormat.Value = 1 Then
  CorrectRR = False
ElseIf Worksheets("RR").Shapes("NewDistu").ControlFormat.Value = 1 Then
  CorrectRR = True
  SpecDistu = False
ElseIf Worksheets("RR").Shapes("SpecDistu").ControlFormat.Value = 1 Then
  CorrectRR = True
  SpecDistu = True
Else
  alert = MsgBox("Error: Please an option for correcting for range restriction", vbCritical, "Choose an option")
  Exit Sub
End If

If Worksheets("RR").Shapes("RRdirect").ControlFormat.Value = 1 Then
  rrDirect = True
  rrIndirect = False
ElseIf Worksheets("RR").Shapes("RRindirect").ControlFormat.Value = 1 Then
  rrDirect = False
  rrIndirect = True
ElseIf CorrectRR Then
  alert = MsgBox("Error: Please choose either direct or indirect range restriction", vbCritical, "Choose an option")
  Exit Sub
End If

If Worksheets("RR").Shapes("ux").ControlFormat.Value = 1 Then
  Observedu = True
ElseIf Worksheets("RR").Shapes("uT").ControlFormat.Value = 1 Then
  Observedu = False
ElseIf CorrectRR And Not SpecDistu Then
  alert = MsgBox("Error: Please choose whether the u values in Column A are for observed scores (ux) or true scores (uT)", vbCritical, "Choose an option")
  Exit Sub
End If


' =====================================
' ===== Get data to meta-analyze ======
' =====================================

' === Correlations and sample sizes ===
k = Application.Count(Worksheets("Correlations").Range("A:A"))
kN = Application.Count(Worksheets("Correlations").Range("B:B"))
ReDim R(k, 2)
For i = 1 To k
  R(i, 1) = Worksheets("Correlations").Cells(i + 1, 1).Value
  R(i, 2) = Worksheets("Correlations").Cells(i + 1, 2).Value
Next i

' Error messages for faulty data
If k = 0 Then
  alert = MsgBox("Error: No correlations entered in Column A of Correlations page", vbCritical, "Missing information")
  Exit Sub
ElseIf k < kN Then
  alert = MsgBox("Error: One or more correlations missing. Please check data entered in Column A of Correlations page.", vbCritical, "Missing information")
  Exit Sub
ElseIf k > kN Then
  alert = MsgBox("Error: One or more sample sizes missing. Please check data entered in Column B of Correlations page.", vbCritical, "Missing information")
  Exit Sub
ElseIf Application.Sum(Worksheets("Correlations").Range("A:A")) > Application.Sum(Worksheets("Correlations").Range("B:B")) Then
  alert = MsgBox("Error: It appears that you have entered correlations and sample sizes in the wrong columns. Please check data entered on the Correlations page.", vbCritical, "Check data")
  Exit Sub
End If

' ========= Reliability of X ==========
' New rxx distribution
If CorrectRxx And Not SpecDistx Then
  ' Get rxx data
  nRxx = Application.Count(Worksheets("rxx").Range("A:A"))
    ReDim RX(nRxx, 2)
    For i = 1 To nRxx
      RX(i, 1) = Worksheets("rxx").Cells(i + 1, 1).Value
      RX(i, 2) = Worksheets("rxx").Cells(i + 1, 2).Value
    Next i
  
  ' Compute rxx distribution
  SumRxx = 0
  SumRxxFreq = 0
  SumRxxSq = 0
  SumQualx = 0
  For i = 1 To nRxx
    SumRxx = SumRxx + RX(i, 1) * RX(i, 2)
    SumRxxSq = SumRxxSq + (RX(i, 1) ^ 2) * RX(1, 2)
    SumQualx = SumQualx + Sqr(RX(i, 1)) * RX(i, 2)
    SumRxxFreq = SumRxxFreq + RX(i, 2)
  Next i
  MeanRxx = SumRxx / SumRxxFreq
  SDRxx = Sqr(WorksheetFunction.Max(0, (SumRxxSq / SumRxxFreq) - (SumRxx / SumRxxFreq) ^ 2))
  MeanQualx = SumQualx / SumRxxFreq
  SDQualx = Sqr(WorksheetFunction.Max(0, (SumRxx / SumRxxFreq) - (SumQualx / SumRxxFreq) ^ 2))

' Prespecified rxx distribution
ElseIf CorrectRxx And SpecDistx Then
  If IsEmpty(Worksheets("rxx").Cells(8, 5)) Then
    noMeanRxx = True
  Else
    noMeanRxx = False
  End If
  If IsEmpty(Worksheets("rxx").Cells(9, 5)) Then
    noSDRxx = True
  Else
    noSDRxx = False
  End If
  If IsEmpty(Worksheets("rxx").Cells(11, 5)) Then
    noMeanQualx = True
  Else
    noMeanQualx = False
  End If
  If IsEmpty(Worksheets("rxx").Cells(12, 5)) Then
    noSDQualx = True
  Else
    noSDQualx = False
  End If
  If noMeanRxx And noSDRxx And noMeanQualx And noSDQualx Then
    alert = MsgBox("Error: Please enter prespecified artifact distribution values for reliability in X", vbCritical, "Missing information")
    Exit Sub
  End If
  ' Mean Rxx
  If noMeanRxx Then
    If noMeanQualx Then
      alert = MsgBox("Error: Please enter mean artifact distribution values for reliability in X", vbCritical, "Missing information")
      Exit Sub
    ElseIf noSDQualx Then
      MeanQualx = Worksheets("rxx").Cells(11, 5).Value
      MeanRxx = MeanQualx * MeanQualx
      FlagRxxOver = "* Mean rxx estimated as mean(" & ChrW(8730) & "rxx) squared because no SD given for " & ChrW(8730) & "rxx. Range restriction estimates are slightly inaccurate."
      flags = flags + 1
    Else
      MeanQualx = Worksheets("rxx").Cells(11, 5).Value
      SDQualx = Worksheets("rxx").Cells(12, 5).Value
      ReDim RqX(1000, 1)
      For i = 1 To 1000
        RqX(i, 1) = TruncNorm(0, 1.1, MeanQualx, SDQualx)
      Next i
      SumRxx = 0
      For i = 1 To 1000
        SumRxx = SumRxx + RqX(i, 1) * RqX(i, 1)
      Next i
      MeanRxx = SumRxx / 1000
    End If
  Else
    MeanRxx = Worksheets("rxx").Cells(8, 5).Value
  End If
  ' SD Rxx
  SDRxx = Worksheets("rxx").Cells(9, 5).Value ' SDRxx is not used in any meta-analysis equation, so it does not need to be estimated if missing.
  ' Mean and SD Qualx
  If noMeanQualx Or noSDQualx Then
    If noSDRxx Then
      If noMeanQualx Then
        MeanQualx = Sqr(MeanRxx)
        FlagQualxUnder = "* Mean " & ChrW(8730) & "rxx estimated as " & ChrW(8730) & "mean(rxx) because no SD given for rxx. Estimated " & ChrW(961) & " is a slight underestimate."
        flags = flags + 1
      Else
        MeanQualx = Worksheets("rxx").Cells(11, 5).Value
      End If
      If noSDQualx Then
        SDQualx = 0
        FlagNoSDQualx = "* No SD given for reliability of X (rxx) or square root of reliability of X. Values of zero assumed. SD" & ChrW(961) & " is an overestimate."
        flags = flags + 1
      Else
        SDQualx = Worksheets("rxx").Cells(12, 5).Value
      End If
    Else
      ReDim RX(1000, 1)
      For i = 1 To 1000
        RX(i, 1) = TruncNorm(0, 1.1, MeanRxx, SDRxx)
      Next i
      SumQualx = 0
      SumRxx = 0
      For i = 1 To 1000
        SumQualx = SumQualx + Sqr(RX(i, 1))
        SumRxx = SumRxx + RX(i, 1)
      Next i
      If noMeanQualx Then
        MeanQualx = SumQualx / 1000
      Else
        MeanQualx = Worksheets("rxx").Cells(11, 5).Value
      End If
      If noSDQualx Then
        SDQualx = Sqr((SumRxx / 1000) - (SumQualx / 1000) ^ 2)
      Else
        SDQualx = Worksheets("rxx").Cells(12, 5).Value
      End If
    End If
  Else
    MeanQualx = Worksheets("rxx").Cells(11, 5).Value
    SDQualx = Worksheets("rxx").Cells(12, 5).Value
  End If
ElseIf Not CorrectRxx Then
  nRxx = 1
  MeanRxx = 1
  MeanQualx = 1
  MeanQualxa = 1
  MeanQualxi = 1
  MeanRxxa = 1
  MeanRxxi = 1
  SDRxx = 0
  SDQualx = 0
  SDRxxa = 0
  SDRxxi = 0
  SDQualxa = 0
  SDQualxi = 0
  ReDim RX(1, 2)
  RX(1, 1) = 1
  RX(1, 2) = 1
End If
  
' ========= Reliability of Y ==========
' New ryy distribution
If CorrectRyy And Not SpecDisty Then
  ' Get ryy data
  nRyy = Application.Count(Worksheets("ryy").Range("A:A"))
    ReDim RY(nRyy, 2)
    For i = 1 To nRyy
      RY(i, 1) = Worksheets("ryy").Cells(i + 1, 1).Value
      RY(i, 2) = Worksheets("ryy").Cells(i + 1, 2).Value
    Next i
    
  ' Compute ryy distribution
  SumRyy = 0
  SumRyyFreq = 0
  SumRyySq = 0
  SumQualy = 0
  For i = 1 To nRyy
    SumRyy = SumRyy + RY(i, 1) * RY(i, 2)
    SumRyySq = SumRyySq + (RY(i, 1) ^ 2) * RY(1, 2)
    SumQualy = SumQualy + Sqr(RY(i, 1)) * RY(i, 2)
    SumRyyFreq = SumRyyFreq + RY(i, 2)
  Next i
  MeanRyy = SumRyy / SumRyyFreq
  SDRyy = Sqr(WorksheetFunction.Max(0, (SumRyySq / SumRyyFreq) - (SumRyy / SumRyyFreq) ^ 2))
  MeanQualy = SumQualy / SumRyyFreq
  SDQualy = Sqr(WorksheetFunction.Max(0, (SumRyy / SumRyyFreq) - (SumQualy / SumRyyFreq) ^ 2))

' Prespecified ryy distribution
ElseIf CorrectRyy And SpecDisty Then
  If IsEmpty(Worksheets("ryy").Cells(8, 5)) Then
    noMeanRyy = True
  Else
    noMeanRyy = False
  End If
  If IsEmpty(Worksheets("ryy").Cells(9, 5)) Then
    noSDRyy = True
  Else
    noSDRyy = False
  End If
  If IsEmpty(Worksheets("ryy").Cells(11, 5)) Then
    noMeanQualy = True
  Else
    noMeanQualy = False
  End If
  If IsEmpty(Worksheets("ryy").Cells(12, 5)) Then
    noSDQualy = True
  Else
    noSDQualy = False
  End If
  If noMeanRyy And noSDRyy And noMeanQualy And noSDQualy Then
    alert = MsgBox("Error: Please enter prespecified artifact distribution values for reliability in Y", vbCritical, "Missing information")
    Exit Sub
  End If
  ' Mean ryy
  MeanRyy = Worksheets("ryy").Cells(8, 5).Value ' Mean Ryy is not used in any meta-analysis equation, so it does not need to be estimated if missing.
  ' SD Ryy
  SDRyy = Worksheets("ryy").Cells(9, 5).Value ' SDRyy is not used in any meta-analysis equation, so it does not need to be estimated if missing.
  ' Mean and SD Qualy
  If noMeanQualy Or noSDQualy Then
    If noSDRyy Then
      If noMeanQualy Then
        MeanQualy = Sqr(MeanRyy)
        FlagQualyUnder = "* Mean " & ChrW(8730) & "ryy estimated as " & ChrW(8730) & "mean(ryy) because no SD given for ryy. Estimated " & ChrW(961) & " is a slight underestimate."
        flags = flags + 1
      Else
        MeanQualy = Worksheets("ryy").Cells(11, 5).Value
      End If
      If noSDQualy Then
        SDQualy = 0
        FlagNoSDQualy = "* No SD given for reliability of Y (ryy) or square root of reliability of Y. Values of zero assumed. SD" & ChrW(961) & " is an overestimate."
        flags = flags + 1
      Else
        SDQualy = Worksheets("ryy").Cells(12, 5).Value
      End If
    Else
      ReDim RY(1000, 1)
      For i = 1 To 1000
        RY(i, 1) = TruncNorm(0, 1.1, MeanRyy, SDRyy)
      Next i
      SumQualy = 0
      SumRyy = 0
      For i = 1 To 1000
        SumQualy = SumQualy + Sqr(RY(i, 1))
        SumRyy = SumRyy + RY(i, 1)
      Next i
      If noMeanQualy Then
        MeanQualy = SumQualy / 1000
      Else
        MeanQualy = Worksheets("ryy").Cells(11, 5).Value
      End If
      If noSDQualy Then
        SDQualy = Sqr((SumRyy / 1000) - (SumQualy / 1000) ^ 2)
      Else
        SDQualy = Worksheets("ryy").Cells(12, 5).Value
      End If
    End If
  Else
    MeanQualy = Worksheets("ryy").Cells(11, 5).Value
    SDQualy = Worksheets("ryy").Cells(12, 5).Value
  End If
ElseIf Not CorrectRyy Then
  nRyy = 1
  MeanRyy = 1
  MeanQualy = 1
  SDRyy = 0
  SDQualy = 0
  ReDim RY(1, 2)
  RY(1, 1) = 1
  RY(1, 2) = 1
End If
  
' ========= Range restriction ==========
' New u distribution
If CorrectRR And Not SpecDistu Then
  ' Get u data
  nU = Application.Count(Worksheets("RR").Range("A:A"))
  ReDim U(nU, 2)
  For i = 1 To nU
    U(i, 1) = Worksheets("RR").Cells(i + 1, 1).Value
    U(i, 2) = Worksheets("RR").Cells(i + 1, 2).Value
  Next i
  ' Compute u distribution
  SumU = 0
  SumUFreq = 0
  SumUSq = 0
  For i = 1 To nU
    SumU = SumU + U(i, 1) * U(i, 2)
    SumUSq = SumUSq + (U(i, 1) ^ 2) * U(1, 2)
    SumUFreq = SumUFreq + U(i, 2)
  Next i
  MeanU = SumU / SumUFreq
  SDu = Sqr(WorksheetFunction.Max(0, (SumUSq / SumUFreq) - (SumU / SumUFreq) ^ 2))
  MeanBigU = 1 / MeanU
  
' Prespecified u distribution
ElseIf CorrectRR And SpecDistu Then
  If IsEmpty(Worksheets("RR").Cells(13, 5)) Then
    noMeanux = True
  Else
    noMeanux = False
  End If
  If IsEmpty(Worksheets("RR").Cells(14, 5)) Then
    noSDux = True
  Else
    noSDux = False
  End If
  If IsEmpty(Worksheets("RR").Cells(16, 5)) Then
    noMeanuT = True
  Else
    noMeanuT = False
  End If
  If IsEmpty(Worksheets("RR").Cells(17, 5)) Then
    noSDuT = True
  Else
    noSDuT = False
  End If
  If noMeanux And noSDux And noMeanuT And noSDQualuT Then
    alert = MsgBox("Error: Please enter prespecified artifact distribution values for range restriction", vbCritical, "Missing information")
    Exit Sub
  End If
  Meanux = Worksheets("RR").Cells(13, 5)
  SDux = Worksheets("RR").Cells(14, 5)
  MeanuT = Worksheets("RR").Cells(16, 5)
  SDuT = Worksheets("RR").Cells(17, 5)
  If Not noMeanux Then MeanBigUx = 1 / Meanux
  If Not noMeanuT Then MeanBigUT = 1 / MeanuT
ElseIf Not CorrectRR Then
  nU = 1
  MeanU = 1
  MeanBigU = 1
  SDu = 0
  ReDim U(1, 2)
  U(1, 1) = 1
  U(1, 2) = 1
End If

' =====================================================
' ===== Transform artifacts for range restriction =====
' =====================================================

' This section of the program computes values for the distributions of rxx_i, rxx_a, ux, and uT if they are not already present

If CorrectRR Then ' Skip this if not correcting for range restriction
  If CorrectRxx Then ' If not correcting for measurement error, then ux and uT are the same
    ' ===== Reliability in X =====
    If Not RelUnrestx Then
      MeanRxxi = MeanRxx
      SDRxxi = SDRxx
      MeanQualxi = MeanQualx
      SDQualxi = SDQualx
      If SpecDistx Then
        If SpecDistu Then
          If noMeanux Then
            Meanux = Sqr((MeanuT ^ 2) / (MeanuT ^ 2 + MeanRxxi * (1 - MeanuT ^ 2)))
            MeanBigUx = 1 / Meanux
          End If
          MeanRxxa = 1 - (Meanux ^ 2 * (1 - MeanRxxi))
          If Not IsEmpty(SDRxxi) Then SDRxxa = Sqr(WorksheetFunction.Max(0, 1 + (Meanux ^ 2 * (Meanux ^ 2 - 2)) + (2 * Meanux ^ 2 * MeanRxxi * (1 - Meanux ^ 2)) + (Meanux ^ 4 * (SDRxxi ^ 2 + MeanRxxi ^ 2)) - MeanRxxa ^ 2))
          MeanQualxa = 1 - (Meanux ^ 2 * (1 - MeanQualxi))
          SDQualxa = Sqr(WorksheetFunction.Max(0, MeanRxxa - MeanQualxa ^ 2))
        ElseIf Not SpecDistu Then
          If Observedu Then
            MeanRxxa = 1 - (MeanU ^ 2 * (1 - MeanRxxi))
            If Not IsEmpty(SDRxxi) Then SDRxxa = Sqr(WorksheetFunction.Max(0, 1 + (MeanU ^ 2 * (MeanU ^ 2 - 2)) + (2 * MeanU ^ 2 * MeanRxxi * (1 - MeanU ^ 2)) + (MeanU ^ 4 * (SDRxxi ^ 2 + MeanRxxi ^ 2)) - MeanRxxa ^ 2))
            MeanQualxa = 1 - (MeanU ^ 2 * (1 - MeanQualxi))
            SDQualxa = Sqr(WorksheetFunction.Max(0, MeanRxxa - MeanQualxa ^ 2))
          ElseIf Not Observedu Then
            Meanux = Sqr((MeanU ^ 2) / (MeanU ^ 2 + MeanRxxi * (1 - MeanU ^ 2)))
            MeanRxxa = 1 - (Meanux ^ 2 * (1 - MeanRxxi))
            If Not IsEmpty(SDRxxi) Then SDRxxa = Sqr(WorksheetFunction.Max(0, 1 + (Meanux ^ 2 * (Meanux ^ 2 - 2)) + (2 * Meanux ^ 2 * MeanRxxi * (1 - Meanux ^ 2)) + (Meanux ^ 4 * (SDRxxi ^ 2 + MeanRxxi ^ 2)) - MeanRxxa ^ 2))
            MeanQualxa = 1 - (Meanux ^ 2 * (1 - MeanQualxi))
            SDQualxa = Sqr(WorksheetFunction.Max(0, MeanRxxa - MeanQualxa ^ 2))
          End If
        End If
      ElseIf Not SpecDistx Then
        ReDim RXi(nRxx, 2)
        RXi = RX
        ReDim RXa(nRxx, 2)
        If SpecDistu Then
          If Not noMeanux Then
            For i = 1 To nRxx
              RXa(i, 1) = 1 - (Meanux ^ 2 * (1 - RXi(i, 1)))
              RXa(i, 2) = RXi(i, 2)
            Next i
          ElseIf noMeanux Then
            Meanux = Sqr((MeanuT ^ 2) / (MeanuT ^ 2 + MeanRxxi * (1 - MeanuT ^ 2)))
            For i = 1 To nRxx
              RXa(i, 1) = 1 - (Meanux ^ 2 * (1 - RXi(i, 1)))
              RXa(i, 2) = RXi(i, 2)
            Next i
          End If
        ElseIf Not SpecDistu Then
          If Observedu Then
            For i = 1 To nRxx
              RXa(i, 1) = 1 - (MeanU ^ 2 * (1 - RXi(i, 1)))
              RXa(i, 2) = RXi(i, 2)
            Next i
          ElseIf Not Observedu Then
            Meanux = Sqr((MeanU ^ 2) / (MeanU ^ 2 + MeanRxxi * (1 - MeanU ^ 2)))
            For i = 1 To nRxx
              RXa(i, 1) = 1 - (Meanux ^ 2 * (1 - RXi(i, 1)))
              RXa(i, 2) = RXi(i, 2)
            Next i
          End If
        End If
        ' Compute new rxx_a distribution values
        SumRxx = 0
        SumRxxSq = 0
        SumQualx = 0
        For i = 1 To nRxx
          SumRxx = SumRxx + RXa(i, 1) * RXa(i, 2)
          SumRxxSq = SumRxxSq + (RXa(i, 1) ^ 2) * RXa(1, 2)
          SumQualx = SumQualx + Sqr(RXa(i, 1)) * RXa(i, 2)
        Next i
        MeanRxxa = SumRxx / SumRxxFreq
        SDRxxa = Sqr(WorksheetFunction.Max(0, (SumRxxSq / SumRxxFreq) - (SumRxx / SumRxxFreq) ^ 2))
        MeanQualxa = SumQualx / SumRxxFreq
        SDQualxa = Sqr(WorksheetFunction.Max(0, (SumRxx / SumRxxFreq) - (SumQualx / SumRxxFreq) ^ 2))
      End If
    ElseIf RelUnrestx Then
      MeanRxxa = MeanRxx
      SDRxxa = SDRxx
      MeanQualxa = MeanQualx
      SDQualxa = SDQualx
      If SpecDistx Then
        If SpecDistu Then
          If noMeanux Then
            Meanux = Sqr((MeanRxxa * MeanuT ^ 2) - MeanRxxa + 1)
            MeanBigUx = 1 / Meanux
          End If
          MeanRxxi = 1 - (MeanBigUx ^ 2 * (1 - MeanRxxa)) ' TODO: Handle the problem with very low ux and rxxa
          If Not IsEmpty(SDRxxa) Then SDRxxi = Sqr(WorksheetFunction.Max(0, 1 + (MeanBigUx ^ 2 * (MeanBigUx ^ 2 - 2)) + (2 * MeanBigUx ^ 2 * MeanRxxa * (1 - MeanBigUx ^ 2)) + (MeanBigUx ^ 4 * (SDRxxa ^ 2 + MeanRxxa ^ 2)) - MeanRxxi ^ 2)) ' TODO: Handle the problem with very low ux and rxxa
          MeanQualxi = 1 - (MeanBigUx ^ 2 * (1 - MeanQualxa)) ' TODO: Handle the problem with very low ux and rxxa
          SDQualxi = Sqr(WorksheetFunction.Max(0, MeanRxxi - MeanQualxi ^ 2))
        ElseIf Not SpecDistu Then
          If Observedu Then
            MeanRxxi = 1 - (MeanBigU ^ 2 * (1 - MeanRxxa)) ' TODO: Handle the problem with very low ux and rxxa
            If Not IsEmpty(SDRxxa) Then SDRxxi = Sqr(1 + (MeanBigU ^ 2 * (MeanBigU ^ 2 - 2)) + (2 * MeanBigU ^ 2 * MeanRxxa * (1 - MeanBigU ^ 2)) + (MeanBigU ^ 4 * (SDRxxa ^ 2 + MeanRxxa ^ 2)) - MeanRxxi ^ 2) ' TODO: Handle the problem with very low ux and rxxa, Protect against floating point errors
            MeanQualxi = 1 - (MeanBigU ^ 2 * (1 - MeanQualxa)) ' TODO: Handle the problem with very low ux and rxxa
            SDQualxi = Sqr(WorksheetFunction.Max(0, MeanRxxi - MeanQualxi ^ 2))
          ElseIf Not Observedu Then
            Meanux = Sqr((MeanRxxa * MeanuT ^ 2) - MeanRxxa + 1)
            MeanBigUx = 1 / Meanux
            MeanRxxi = 1 - (MeanBigUx ^ 2 * (1 - MeanRxxa)) ' TODO: Handle the problem with very low ux and rxxa
            If Not IsEmpty(SDRxxa) Then SDRxxi = Sqr(WorksheetFunction.Max(0, 1 + (MeanBigUx ^ 2 * (MeanBigUx ^ 2 - 2)) + (2 * MeanBigUx ^ 2 * MeanRxxa * (1 - MeanBigUx ^ 2)) + (MeanBigUx ^ 4 * (SDRxxa ^ 2 + MeanRxxa ^ 2)) - MeanRxxi ^ 2)) ' Handle the problem with very low ux and rxxa
            MeanQualxi = 1 - (MeanBigUx ^ 2 * (1 - MeanQualxa)) ' TODO: Handle the problem with very low ux and rxxa
            SDQualxi = Sqr(WorksheetFunction.Max(0, MeanRxxi - MeanQualxi ^ 2))
          End If
        End If
      ElseIf Not SpecDistx Then
        ReDim RXa(nRxx, 2)
        RXa = RX
        ReDim RXi(nRxx, 2)
        If SpecDistu Then
          If Not noMeanux Then
            For i = 1 To nRxx
              RXi(i, 1) = 1 - (MeanBigUx ^ 2 * (1 - RXa(i, 1))) ' TODO: Handle the problem with very low ux and rxxa
              RXi(i, 2) = RXa(i, 2)
            Next i
          ElseIf noMeanux Then
            Meanux = Sqr((MeanRxxa * MeanuT ^ 2) - MeanRxxa + 1)
            MeanBigUx = 1 / Meanux
            For i = 1 To nRxx
              RXi(i, 1) = 1 - (MeanBigUx ^ 2 * (1 - RXa(i, 1))) ' TODO: Handle the problem with very low ux and rxxa
              RXi(i, 2) = RXa(i, 2)
            Next i
          End If
        ElseIf Not SpecDistu Then
          If Observedu Then
            For i = 1 To nRxx
              RXi(i, 1) = 1 - (MeanBigU ^ 2 * (1 - RXa(i, 1))) ' TODO: Handle the problem with very low ux and rxxa
              RXi(i, 2) = RXa(i, 2)
            Next i
          ElseIf Not Observedu Then
            Meanux = Sqr((MeanRxxa * MeanuT ^ 2) - MeanRxxa + 1)
            MeanBigUx = 1 / Meanux
            For i = 1 To nRxx
              RXi(i, 1) = 1 - (MeanBigUx ^ 2 * (1 - RXa(i, 1))) ' TODO: Handle the problem with very low ux and rxxa
              RXi(i, 2) = RXa(i, 2)
            Next i
          End If
        End If
        ' Compute new rxx_i distribution values
        SumRxx = 0
        SumRxxSq = 0
        SumQualx = 0
        For i = 1 To nRxx
          SumRxx = SumRxx + RXi(i, 1) * RXi(i, 2)
          SumRxxSq = SumRxxSq + (RXi(i, 1) ^ 2) * RXi(1, 2)
          SumQualx = SumQualx + Sqr(RXi(i, 1)) * RXi(i, 2)
        Next i
        MeanRxxi = SumRxx / SumRxxFreq
        SDRxxi = Sqr(WorksheetFunction.Max(0, (SumRxxSq / SumRxxFreq) - (SumRxx / SumRxxFreq) ^ 2))
        MeanQualxi = SumQualx / SumRxxFreq
        SDQualxi = Sqr(WorksheetFunction.Max(0, (SumRxx / SumRxxFreq) - (SumQualx / SumRxxFreq) ^ 2))
      End If
    End If
    ' ==== Range restriction values ====
    If SpecDistu Then
      If noMeanux Then Meanux = Sqr((MeanuT ^ 2) / (MeanuT ^ 2 + MeanRxxi * (1 - MeanuT ^ 2)))
      If noSDux Then SDux = Sqr(SDuT ^ 2 * MeanRxxa)
      If noMeanuT Then MeanuT = Sqr(((Meanux ^ 2) - (1 - MeanRxxa)) / MeanRxxa) ' TODO: Handle the problem with very low ux and rxxa
      If noSDuT Then SDuT = Sqr(SDux ^ 2 / MeanRxxa)
      MeanBigUx = 1 / Meanux
      MeanBigUT = 1 / MeanuT
    ElseIf Not SpecDistu Then
      If Observedu Then
        Meanux = MeanU
        SDux = SDu
        MeanBigUx = MeanBigU
        ReDim Ux(nU, 2)
        Ux = U
        ReDim UT(nU, 2)
        For i = 1 To nU
          UT(i, 1) = Sqr(((Ux(i, 1) ^ 2) - (1 - MeanRxxa)) / MeanRxxa) ' TODO: Handle the problem with very low ux and rxxa
          UT(i, 2) = Ux(i, 2)
        Next i
        ' Compute new ux distribution values
        SumU = 0
        SumUSq = 0
        For i = 1 To nU
          SumU = SumU + UT(i, 1) * UT(i, 2)
          SumUSq = SumUSq + (UT(i, 1) ^ 2) * UT(1, 2)
        Next i
        MeanuT = SumU / SumUFreq
        SDuT = Sqr(WorksheetFunction.Max(0, (SumUSq / SumUFreq) - (SumU / SumUFreq) ^ 2))
        MeanBigUT = 1 / MeanuT
      ElseIf Not Observedu Then
        MeanuT = MeanU
        SDuT = SDu
        MeanBigUT = MeanBigU
        ReDim UT(nU, 2)
        UT = U
        ReDim Ux(nU, 2)
        For i = 1 To nU
          Ux(i, 1) = Sqr((UT(i, 1) ^ 2) / (UT(i, 1) ^ 2 + MeanRxxi * (1 - UT(i, 1) ^ 2)))
          Ux(i, 2) = UT(i, 2)
        Next i
        ' Compute new ux distribution values
        SumU = 0
        SumUSq = 0
        For i = 1 To nU
          SumU = SumU + Ux(i, 1) * Ux(i, 2)
          SumUSq = SumUSq + (Ux(i, 1) ^ 2) * Ux(1, 2)
        Next i
        Meanux = SumU / SumUFreq
        SDux = Sqr(WorksheetFunction.Max(0, (SumUSq / SumUFreq) - (SumU / SumUFreq) ^ 2))
        MeanBigUx = 1 / Meanux
      End If
    End If
  Else
    ' If there is no measurement error in X, then distributions of ux and uT are the same
    If SpecDistu Then
      If noMeanux Then
        Meanux = MeanuT
        MeanBigUx = 1 / Meanux
      End If
      If noSDux Then SDux = SDuT
      If noMeanuT Then
        MeanuT = Meanux
        MeanBigUT = 1 / MeanuT
      End If
      If noSDuT Then SDuT = SDux
    Else
      Meanux = MeanU
      SDux = SDu
      MeanBigUx = MeanBigU
      MeanuT = MeanU
      SDuT = SDu
      MeanBigUT = MeanBigU
      ReDim Ux(nU, 2)
      ReDim UT(nU, 2)
      Ux = U
      UT = U
    End If
  End If
End If

' =====================================
' ===== Bare bones meta-analysis ======
' =====================================

' Compute mean uncorrected r
Ntotal = 0
sumR = 0
For i = 1 To k
  sumR = sumR + R(i, 2) * R(i, 1)
  Ntotal = Ntotal + R(i, 2)
Next i
meanR = sumR / Ntotal
meanN = Ntotal / k

' Correct r values for small sample size bias
aR = 1 - (1 - meanR ^ 2) / ((2 * meanN) - 2)
meanR = meanR / aR

' Compute expected sampling error variance of observed r's
unexpVar = 1 - (meanR ^ 2)
SampErrVar = (unexpVar ^ 2) / (meanN - 1)

' Compute observed variance of observed r's
ObsSSQ = 0
For i = 1 To k
  ObsSSQ = ObsSSQ + R(i, 2) * (R(i, 1) - meanR) ^ 2
Next i
ObsVar = ObsSSQ / Ntotal
SDobs = Sqr(ObsVar)

' Compute percent of variance due to sampling error
If ObsVar < 1E-16 Then
  PerVarSamp = "No Obs. Var."
Else
  PerVarSamp = (SampErrVar / ObsVar)
End If

' =====================================
' ===== Apply artifact corrections ====
' =====================================

' TODO: Implement Le et al. (2013) range restriction correction (p. 193 H&S 3ed)
' TODO: Implement Alexander et al. (1987) double range restriction correction (p. 193 H&S 3ed)
' TODO: Consider Raju, Burke, Normand (1991) methods for accounting for sampling error in artifacts

' === Compute true score mean r (rho) ===
If Not CorrectRR Then ' No range restriction
  Rho = meanR / (MeanQualy * MeanQualx)
  RhoValidity = meanR / MeanQualy
ElseIf rrDirect Then ' Direct range restriction
  RhoValidityRestricted = meanR / MeanQualy
  RhoValidity = RhoValidityRestricted * MeanBigUx / Sqr(((MeanBigUx ^ 2) - 1) * (RhoValidityRestricted ^ 2) + 1)
  Rho = RhoValidity / MeanQualxa
ElseIf rrIndirect Then ' Indirect range restriction
  RhoRestricted = meanR / (MeanQualy * MeanQualxi)
  Rho = RhoRestricted * MeanBigUT / Sqr(((MeanBigUT ^ 2) - 1) * (RhoRestricted ^ 2) + 1)
  RhoValidity = Rho * MeanQualxa
End If

' === Compute variance due to artifact differences ===
If Not CorrectRR Then ' No range restriction
  If (Not CorrectRxx Or Not SpecDistx) And (Not CorrectRyy Or Not SpecDisty) Then
  ' If all distributions are new, then use interactive method to estimate variance
    Taylor = False
    SumRatten = 0
    SumRSQatten = 0
    SumArtFreq = 0
    For IXX = 1 To nRxx
      Qualx = Sqr(RX(IXX, 1))
      For ICC = 1 To nRyy
        Qualy = Sqr(RY(ICC, 1))
        Ratten = Qualx * Qualy * Rho
        ArtFreq = RX(IXX, 2) * RY(ICC, 2)
        SumRatten = SumRatten + ArtFreq * Ratten
        SumRSQatten = SumRSQatten + ArtFreq * Ratten * Ratten
        SumArtFreq = SumArtFreq + ArtFreq
      Next ICC
    Next IXX
    SumRatten = SumRatten / SumArtFreq
    ArtVar = SumRSQatten / SumArtFreq - SumRatten * SumRatten
  Else
  ' If some of the distributions are pre-specified, then use Raju and Burke's (1983) Taylor Series 2 model to estimate variance
  ' Note that R&B1983 had errors in their formulas and multipled F, G by .5
    Taylor = True
    E = meanR / Rho
    F = (meanR / MeanQualx)
    G = (meanR / MeanQualy)
    VarRho = (ObsVar - SampErrVar - (F ^ 2) * (SDQualx ^ 2) - (G ^ 2) * (SDQualy ^ 2)) / (E ^ 2) / (aR ^ 2) ' aR is correcting for the disattenuation of the slight bias in the sample correlation coefficient
    SDrho = Sqr(WorksheetFunction.Max(0, VarRho))
    SDrhoValidity = SDrho * MeanQualx
    ResVar = VarRho * ((MeanQualx * MeanQualy) ^ 2)
    SDres = SDrho * (MeanQualx * MeanQualy)
    PredVar = ObsVar - ResVar
    SDpred = Sqr(PredVar)
    
    ' Alternative method:
      ' Multiplicative (noninteractive) model for estimating true variance
      ' Results are virtually identical to the Raju and Burke method for these artifacts
      ' Taylor = False
      ' A = MeanQualx * MeanQualy
      ' V = (SDQualx/MeanQualx)^2 + (SDQualy/MeanQualy)^2
      ' ArtVar = Rho^2 * A^2 * V
      ' Then use nonlinear function below to estimate VarRho (or just divide by product of MeanRxx and MeanRyy)
  End If
ElseIf rrDirect Then ' Direct range restriction
  If (Not CorrectRxx Or Not SpecDistx) And (Not CorrectRyy Or Not SpecDisty) And Not SpecDistu Then
  ' If all distributions are new, then use interactive method to estimate variance
      Taylor = False
      SumRatten = 0
      SumRSQatten = 0
      SumArtFreq = 0
      For IXX = 1 To nRxx
        Qualx = Sqr(RXa(IXX, 1))
        For ICC = 1 To nRyy
          Qualy = Sqr(RY(ICC, 1))
          For IU = 1 To nU
            uval = Ux(IU, 1)
            Ratten = Qualx * Rho
            RRatten = ((uval ^ 2) - 1) * (Ratten ^ 2) + 1
            Ratten = uval * Ratten / Sqr(RRatten)
            Ratten = Ratten * Qualy
            ArtFreq = RXa(IXX, 2) * RY(ICC, 2) * Ux(IU, 2)
            SumRatten = SumRatten + ArtFreq * Ratten
            SumRSQatten = SumRSQatten + ArtFreq * Ratten * Ratten
            SumArtFreq = SumArtFreq + ArtFreq
          Next IU
        Next ICC
      Next IXX
      SumRatten = SumRatten / SumArtFreq
      ArtVar = SumRSQatten / SumArtFreq - SumRatten * SumRatten
  Else
  ' If some of the distributions are pre-specified, then use Raju and Burke's (1983) Taylor Series 2 model to estimate variance
  ' Note that R&B1983 had errors in their formulas and multipled F, G by .5
    Taylor = True
    E = (meanR / Rho) + (((meanR ^ 3) * (1 - (Meanux ^ 2))) / (Rho * (Meanux ^ 2)))
    F = ((meanR / MeanQualxa) + (((meanR ^ 3) * (1 - (Meanux ^ 2))) / (MeanQualxa * (Meanux ^ 2))))
    G = ((meanR / MeanQualy) + (((meanR ^ 3) * (1 - (Meanux ^ 2))) / (MeanQualy * (Meanux ^ 2))))
    H = (meanR - (meanR ^ 3)) / Meanux
    VarRho = (ObsVar - SampErrVar - (F ^ 2) * (SDQualxa ^ 2) - (G ^ 2) * (SDQualy ^ 2) - (H ^ 2) * (SDux ^ 2)) / (E ^ 2) / (aR ^ 2) ' aR is correcting for the disattenuation of the slight bias in the sample correlation coefficient
    SDrho = Sqr(WorksheetFunction.Max(0, VarRho))
    SDrhoValidity = SDrho * MeanQualxa
    ' Estimate residual distribution of r using reverse of of Law et al.'s (1994) non-linear procedure
    FR = Array(0, 0.0004, 0.0006, 0.0008, 0.001, 0.0014, 0.0018, 0.0022, 0.0028, 0.0036, 0.0044, 0.0054, 0.0066, 0.0079, 0.0094, 0.0111, 0.013, 0.015, 0.0171, 0.0194, 0.0218, 0.0242, 0.0266, 0.029, 0.0312, 0.0333, 0.0352, 0.0368, 0.0381, 0.0391, 0.0397, 0.0399, 0.0397, 0.0391, 0.0381, 0.0368, 0.0352, 0.0333, 0.0312, 0.029, 0.0266, 0.0242, 0.0218, 0.0194, 0.0171, 0.015, 0.013, 0.0111, 0.0094, 0.0079, 0.0066, 0.0054, 0.0044, 0.0036, 0.0028, 0.0022, 0.0018, 0.0014, 0.001, 0.0008, 0.0006, 0.0004)
    sumRnonlinatten = 0
    SSQRnonlinatten = 0
    SumDist = 0
    For i = 1 To 61
      rDist = Rho + (i - 31) * 0.1 * SDrho
      rDist = rDist * MeanQualxa
      rDist = rDist * (Meanux / Sqr(1 + ((Meanux ^ 2) - 1) * (rDist ^ 2)))
      rDist = rDist * MeanQualy
      sumRnonlinatten = sumRnonlinatten + FR(i) * rDist
      SSQRnonlinatten = SSQRnonlinatten + FR(i) * (rDist ^ 2)
      SumDist = SumDist + FR(i)
    Next i
    meanRnonlinatten = sumRnonlinatten / SumDist
    If meanR = meanRnonlinatten Then
      NonLinCheck = True
    Else
      NonLinCheck = False
    End If
    ResVar = SSQRnonlinatten / SumDist - (meanRnonlinatten ^ 2)
    SDres = Sqr(WorksheetFunction.Max(0, ResVar))
    PredVar = ObsVar - ResVar
    SDpred = Sqr(PredVar)
    
    ' Alternative method:
      ' Multiplicative (noninteractive) model for estimating true variance
      ' Results for the Raju and Burke method are more accurate for these artifacts because they don't assume that c is independent of Qualx and Qualy
      ' Taylor = False
      ' c = Sqr( Meanux^2 + (MeanR^2) * (1 - Meanux^2) )
      ' SDc = Sqr( MeanR^2 - c^2 + ( (1 - MeanR^2) * (Meanux^2 + SDux^2) ) )
      ' A = MeanQualx * MeanQualy * c
      ' V = (SDQualx/MeanQualx)^2 + (SDQualy/MeanQualy)^2 + (SDc/c)^2
      ' ArtVar = Rho^2 * A^2 * V
      ' Then use nonlinear function below to estimate VarRho
    End If
ElseIf rrIndirect Then ' Indirect range restriction
  If (Not CorrectRxx Or Not SpecDistx) And (Not CorrectRyy Or Not SpecDisty) And Not SpecDistu Then
  ' If all distributions are new, then use interactive method to estimate variance
      Taylor = False
      SumRatten = 0
      SumRSQatten = 0
      SumArtFreq = 0
      For IXX = 1 To nRxx
        Qualx = Sqr(RXi(IXX, 1))
        For ICC = 1 To nRyy
          Qualy = Sqr(RY(ICC, 1))
          For IU = 1 To nU
            uval = UT(IU, 1)
            RRatten = Rho * (uval / Sqr((uval ^ 2) * (Rho ^ 2) + 1 - (Rho ^ 2)))
            Ratten = RRatten * Qualx * Qualy
            ArtFreq = RXi(IXX, 2) * RY(ICC, 2) * UT(IU, 2)
            SumRatten = SumRatten + ArtFreq * Ratten
            SumRSQatten = SumRSQatten + ArtFreq * Ratten * Ratten
            SumArtFreq = SumArtFreq + ArtFreq
          Next IU
        Next ICC
      Next IXX
      SumRatten = SumRatten / SumArtFreq
      ArtVar = SumRSQatten / SumArtFreq - SumRatten * SumRatten
  Else
  ' If some of the distributions are pre-specified, then use the Hunter et al. (2006) Taylor Series Method
    Taylor = True
    A = 1 / Sqr((MeanuT ^ 2) * (Rho ^ 2) - (Rho ^ 2) + 1)
    B = 1 / Sqr((MeanuT ^ 2) * (MeanQualxa ^ 2) - (MeanQualxa ^ 2) + 1)
    b1 = (meanR / MeanQualxa) - meanR * MeanQualxa * (B ^ 2) * ((MeanuT ^ 2) - 1)
    b2 = meanR / MeanQualy
    b3 = (2 * meanR / MeanuT) - (meanR * MeanuT * (MeanQualxa ^ 2) * (B ^ 2)) - (meanR * MeanuT * (Rho ^ 2) * (A ^ 2))
    b4 = (meanR / Rho) - meanR * Rho * (A ^ 2) * ((MeanuT ^ 2) - 1)
    VarRho = (ObsVar - ((b1 ^ 2) * (SDQualxa ^ 2) + (b2 ^ 2) * (SDQualy ^ 2) + (b3 ^ 2) * (SDuT ^ 2) + SampErrVar)) / (b4 ^ 2) / (aR ^ 2) ' aR is correcting for the disattenuation of the slight bias in the sample correlation coefficient
    SDrho = Sqr(WorksheetFunction.Max(0, VarRho))
    SDrhoValidity = SDrho * MeanQualxa
    
  ' Estimate residual distribution of r using reverse of of Law et al.'s (1994) non-linear procedure
    FR = Array(0, 0.0004, 0.0006, 0.0008, 0.001, 0.0014, 0.0018, 0.0022, 0.0028, 0.0036, 0.0044, 0.0054, 0.0066, 0.0079, 0.0094, 0.0111, 0.013, 0.015, 0.0171, 0.0194, 0.0218, 0.0242, 0.0266, 0.029, 0.0312, 0.0333, 0.0352, 0.0368, 0.0381, 0.0391, 0.0397, 0.0399, 0.0397, 0.0391, 0.0381, 0.0368, 0.0352, 0.0333, 0.0312, 0.029, 0.0266, 0.0242, 0.0218, 0.0194, 0.0171, 0.015, 0.013, 0.0111, 0.0094, 0.0079, 0.0066, 0.0054, 0.0044, 0.0036, 0.0028, 0.0022, 0.0018, 0.0014, 0.001, 0.0008, 0.0006, 0.0004)
    sumRnonlinatten = 0
    SSQRnonlinatten = 0
    SumDist = 0
    For i = 1 To 61
      rDist = Rho + (i - 31) * 0.1 * SDrho
      rDist = rDist * (MeanuT / Sqr((MeanuT ^ 2) * (rDist ^ 2) + 1 - (rDist ^ 2))) * MeanQualxi * MeanQualy
      sumRnonlinatten = sumRnonlinatten + FR(i) * rDist
      SSQRnonlinatten = SSQRnonlinatten + FR(i) * (rDist ^ 2)
      SumDist = SumDist + FR(i)
    Next i
    meanRnonlinatten = sumRnonlinatten / SumDist
    If meanR = meanRnonlinatten Then
      NonLinCheck = True
    Else
      NonLinCheck = False
    End If
    ResVar = SSQRnonlinatten / SumDist - (meanRnonlinatten ^ 2)
    SDres = Sqr(WorksheetFunction.Max(0, ResVar))
    PredVar = ObsVar - ResVar
    SDpred = Sqr(PredVar)
  End If
End If

' ======================================================
' ===== Compute residual distributions of r and rho ====
' ======================================================

' === For all non-Taylor Series methods, estimate residual distribution for obs. r, ===
' === then estimate rho distribution using Law et al.'s (1994) non-linear procedure ===
If Not Taylor Then
  ' Compute residual variance and SD
  ResVar = (ObsVar - SampErrVar - ArtVar) / (aR ^ 2) ' aR is correcting for the disattenuation of the slight bias in the sample correlation coefficient
  SDres = Sqr(WorksheetFunction.Max(0, ResVar))

  ' Computed predicted variance and SD
  PredVar = SampErrVar + ArtVar
  SDpred = Sqr(PredVar)

  ' === Compute SDrho using non-linear procedure from Law et al., 1994, (Hunter & Schmidt, 2004, p. 199) ==
  FR = Array(0, 0.0004, 0.0006, 0.0008, 0.001, 0.0014, 0.0018, 0.0022, 0.0028, 0.0036, 0.0044, 0.0054, 0.0066, 0.0079, 0.0094, 0.0111, 0.013, 0.015, 0.0171, 0.0194, 0.0218, 0.0242, 0.0266, 0.029, 0.0312, 0.0333, 0.0352, 0.0368, 0.0381, 0.0391, 0.0397, 0.0399, 0.0397, 0.0391, 0.0381, 0.0368, 0.0352, 0.0333, 0.0312, 0.029, 0.0266, 0.0242, 0.0218, 0.0194, 0.0171, 0.015, 0.013, 0.0111, 0.0094, 0.0079, 0.0066, 0.0054, 0.0044, 0.0036, 0.0028, 0.0022, 0.0018, 0.0014, 0.001, 0.0008, 0.0006, 0.0004)
  sumRnonlincorr = 0
  SSQRnonlincorr = 0
  SumDist = 0
  For i = 1 To 61
    rDist = meanR + (i - 31) * 0.1 * SDres
    If Not CorrectRR Then ' No range restriction
      rDist = rDist / (MeanQualy * MeanQualx)
    ElseIf rrDirect Then ' Direct range restriction
      rDist = rDist / MeanQualy
      rDist = rDist * MeanBigUx / Sqr(((MeanBigUx ^ 2) - 1) * (rDist ^ 2) + 1)
      rDist = rDist / MeanQualxa
    ElseIf rrIndirect Then ' Indirect range restriction
      rDist = rDist / (MeanQualy * MeanQualxi)
      rDist = rDist * MeanBigUT / Sqr(((MeanBigUT ^ 2) - 1) * (rDist ^ 2) + 1)
    End If
    sumRnonlincorr = sumRnonlincorr + FR(i) * rDist
    SSQRnonlincorr = SSQRnonlincorr + FR(i) * (rDist ^ 2)
    SumDist = SumDist + FR(i)
  Next i
  Rho = sumRnonlincorr / SumDist
  VarRho = SSQRnonlincorr / SumDist - (Rho ^ 2)
  SDrho = Sqr(WorksheetFunction.Max(0, VarRho))
  If Not CorrectRR Then ' No range restriction
    RhoValidity = Rho * MeanQualx
    SDrhoValidity = SDrho * MeanQualx
  Else ' Either direct or indirect range restriction
    RhoValidity = Rho * MeanQualxa
    SDrhoValidity = SDrho * MeanQualxa
  End If
End If

' Compute percent variance accounted for
If ObsVar < 1E-16 Then
  PerVarAcc = "No Obs. Var."
Else
  PerVarAcc = PredVar / ObsVar
End If

' ========================================================
' ===== Compute confidence and credibility intervals =====
' ========================================================

' Confidence Interval - 90% (meanR)
SEmeanR = SDobs / Sqr(k) / aR ' aR is correcting for the disattenuation of the slight bias in the sample correlation coefficient
If k < 30 Then
  df = k - 1
  If CompatMode Then
    Crit = Application.TInv(0.1, df)
  Else
    Crit = Application.T_Inv_2T(0.1, df)
  End If
  UpCImeanR = meanR + Crit * SEmeanR
  LoCImeanR = meanR - Crit * SEmeanR
Else
  UpCImeanR = meanR + 1.64 * SEmeanR
  LoCImeanR = meanR - 1.64 * SEmeanR
End If

' Confidence Interval - 90%  (True validity)
If Not CorrectRR Then ' No range restriction
  UpCIrhoValidity = UpCImeanR / MeanQualy
  LoCIrhoValidity = LoCImeanR / MeanQualy
ElseIf rrDirect Then ' Direct range restriction
  UpCIrhoValidity = (UpCImeanR / MeanQualy) * MeanBigUx / Sqr(((MeanBigUx ^ 2) - 1) * ((UpCImeanR / MeanQualy) ^ 2) + 1)
  LoCIrhoValidity = (LoCImeanR / MeanQualy) * MeanBigUx / Sqr(((MeanBigUx ^ 2) - 1) * ((LoCImeanR / MeanQualy) ^ 2) + 1)
ElseIf rrIndirect Then ' Indirect range restriction
  UpCIrhoValidity = ((UpCImeanR / (MeanQualy * MeanQualxi)) * MeanBigUT / Sqr(((MeanBigUT ^ 2) - 1) * ((UpCImeanR / (MeanQualy * MeanQualxi)) ^ 2) + 1)) * MeanQualxa
  LoCIrhoValidity = ((LoCImeanR / (MeanQualy * MeanQualxi)) * MeanBigUT / Sqr(((MeanBigUT ^ 2) - 1) * ((LoCImeanR / (MeanQualy * MeanQualxi)) ^ 2) + 1)) * MeanQualxa
End If

' Confidence Interval - 90%  (Rho)
If Not CorrectRR Then ' No range restriction
  UpCIrho = UpCIrhoValidity / MeanQualx
  LoCIrho = LoCIrhoValidity / MeanQualx
Else ' Either direct or indirect range restriction
  UpCIrho = UpCIrhoValidity / MeanQualxa
  LoCIrho = LoCIrhoValidity / MeanQualxa
End If

' Credibility Interval - 80% (meanR)
UpCVmeanR = meanR + 1.28 * SDres
LoCVmeanR = meanR - 1.28 * SDres

' Credibility Interval - 80% (True validity)
UpCVrhoValidity = RhoValidity + 1.28 * SDrhoValidity
LoCVrhoValidity = RhoValidity - 1.28 * SDrhoValidity

' Credibility Interval - 80% (Rho)
UpCVrho = Rho + 1.28 * SDrho
LoCVrho = Rho - 1.28 * SDrho

' TODO: Add options to choose different widths of intervals

If SpecDisty Then SumRyyFreq = "Pre-specified"
If SpecDistx Then SumRxxFreq = "Pre-specified"
If SpecDistu Then SumUFreq = "Pre-specified"

If Not CorrectRyy Then SumRyyFreq = "Not corrected"
If Not CorrectRxx Then SumRxxFreq = "Not corrected"
If Not CorrectRR Then SumUFreq = "Not corrected"

' Fail-safe k and N
If IsEmpty(Worksheets("Correlations").Range("F17")) Then
  rc = 0.05
Else
  rc = Worksheets("Correlations").Cells(17, 6).Value
End If
If Not IsNumeric(rc) Then
  alert = MsgBox("Error: Fail-safe threshold value must be a number.", vbCritical, "Check data")
  Exit Sub
ElseIf rc > 1 Or rc < -1 Then
    alert = MsgBox("Error: Fail-safe threshold value must be valid correlation (between -1 and 1).", vbCritical, "Check data")
  Exit Sub
End If
If IsEmpty(Worksheets("Correlations").Range("F19")) Then
  rFS = 0
Else
  rFS = Worksheets("Correlations").Cells(19, 6).Value
End If
If Not IsNumeric(rFS) Then
  alert = MsgBox("Error: Fail-safe mean file drawer value must be a number.", vbCritical, "Check data")
  Exit Sub
ElseIf rFS > 1 Or rFS < -1 Then
    alert = MsgBox("Error: Fail-safe mean file drawer value must be valid correlation (between -1 and 1).", vbCritical, "Check data")
  Exit Sub
End If
kFS = (k * (meanR - rc)) / (rc - rFS)
NFS = kFS * meanN

' ==========================
' ===== Output results =====
' ==========================

Worksheets("Output").Cells.ClearContents
Worksheets("Output").Cells.ClearFormats

' === Build results array ===

' Header
ReDim Main(10, 10)
Main(0, 0) = "Meta-analysis Results"
Main(1, 0) = "Correlations"
Main(2, 0) = "Corrected using artifact distributions"
If flags > 0 Then
  Main(0, 2) = "Warnings were generated. See the 'Alerts' tab for more information."
End If

' Model estimation parameters
Main(1, 2) = "Weights"
Main(2, 2) = "Sample size"
Main(1, 3) = "rxx"
Main(1, 4) = "ryy"
Main(1, 5) = "Range restriction in X"
If Taylor Then
  Main(1, 7) = "True variance estimated using Taylor Series approximation model"
Else
  Main(2, 7) = "True variance estimated using nonlinear interactive model"
End If
If CorrectRxx Then
  Main(2, 3) = "Corrected"
Else
  Main(2, 3) = "Not corrected"
End If
If CorrectRyy Then
  Main(2, 4) = "Corrected"
Else
  Main(2, 4) = "Not corrected"
End If
If Not CorrectRR Then
  Main(2, 5) = "Not corrected"
ElseIf rrDirect Then
  Main(2, 5) = "Corrected for direct RR"
ElseIf rrIndirect Then
  Main(2, 5) = "Corrected for indirect RR"
End If

' Main results table
Main(4, 0) = "Main Results"
Main(5, 0) = "Recommended results table"
Main(6, 0) = "True score correlations"
Main(7, 0) = "Validity generalization"
Main(5, 1) = "N (Total sample)"
Main(6, 1) = Ntotal
Main(7, 1) = Ntotal
Main(5, 2) = "k (No. correlations)"
Main(6, 2) = k
Main(7, 2) = k
Main(5, 3) = "Mean Uncorrected r"
Main(6, 3) = meanR
Main(7, 3) = meanR
Main(5, 4) = "Observed SDr"
Main(6, 4) = SDobs
Main(7, 4) = SDobs
Main(5, 5) = "Residual SDr (SDres)"
Main(6, 5) = SDres
Main(7, 5) = SDres
Main(5, 6) = "SE of Mean Uncorrected r"
Main(6, 6) = SEmeanR
Main(7, 6) = SEmeanR
Main(5, 7) = "Mean Corrected " & ChrW(961)
Main(6, 7) = Rho
Main(7, 7) = RhoValidity
Main(5, 8) = "SD" & ChrW(961)
Main(6, 8) = SDrho
Main(7, 8) = SDrhoValidity
Main(5, 9) = "90% Conf. Int. (" & ChrW(961) & ")"
Main(6, 9) = Format(LoCIrho, ".00") & ", " & Format(UpCIrho, ".00")
Main(7, 9) = Format(LoCIrhoValidity, ".00") & ", " & Format(UpCIrhoValidity, ".00")
Main(5, 10) = "80% Cred. Int. (" & ChrW(961) & ")"
Main(6, 10) = Format(LoCVrho, ".00") & ", " & Format(UpCVrho, ".00")
Main(7, 10) = Format(LoCVrhoValidity, ".00") & ", " & Format(UpCVrhoValidity, ".00")

' Artifacts table header
Main(9, 0) = "Artifact Distribution"
Main(10, 0) = "Recommended results table"
Main(10, 1) = "No. artifact values"
Main(10, 2) = "Mean artifact value"
Main(10, 3) = "SD of artifact values"
Main(10, 4) = "Mean SQRT of reliability"
Main(10, 5) = "SD of SQRT of reliability"

' Output and format header, model estimation parameters, main results table, artifacts table header
Worksheets("Output").Range("A1:K11") = Main
Worksheets("Output").Range("A1, C1, C2:H2, A5, A10").Font.Bold = True
Worksheets("Output").Range("C1").Font.Color = vbRed
Worksheets("Output").Range("B7:C8").NumberFormat = "0"
Worksheets("Output").Range("D7:I8").NumberFormat = ".00"
Worksheets("Output").Range("B6:K6,B11:F11").WrapText = True

' Artifact distributions
If Not CorrectRR Then
  ' Prepare output array
  ReDim Artifacts(2, 5)
  Artifacts(0, 0) = "rxx"
  Artifacts(1, 0) = "ryy"
  Artifacts(2, 0) = "Range restriction (u)"
  Artifacts(0, 1) = SumRxxFreq
  Artifacts(1, 1) = SumRyyFreq
  Artifacts(2, 1) = SumUFreq
  If CorrectRxx Then
    Artifacts(0, 2) = MeanRxx
    Artifacts(0, 3) = SDRxx
    Artifacts(0, 4) = MeanQualx
    Artifacts(0, 5) = SDQualx
  Else
    Artifacts(0, 2) = "--"
    Artifacts(0, 3) = "--"
    Artifacts(0, 4) = "--"
    Artifacts(0, 5) = "--"
  End If
  If CorrectRyy Then
    Artifacts(1, 2) = MeanRyy
    Artifacts(1, 3) = SDRyy
    Artifacts(1, 4) = MeanQualy
    Artifacts(1, 5) = SDQualy
  Else
    Artifacts(1, 2) = "--"
    Artifacts(1, 3) = "--"
    Artifacts(1, 4) = "--"
    Artifacts(1, 5) = "--"
  End If
  Artifacts(2, 2) = "--"
  Artifacts(2, 3) = "--"
  
  ' Output and format artifacts table
  Worksheets("Output").Range("A12:F14") = Artifacts
  Worksheets("Output").Range("B12:B14").NumberFormat = "0"
  Worksheets("Output").Range("C12:F14").NumberFormat = ".00"
Else
  ReDim Artifacts(5, 5)
  Artifacts(0, 0) = "Reliability of X (Unrestricted population: rxx_a)"
  Artifacts(1, 0) = "Reliability of Y (Restricted population: ryy_i)"
  Artifacts(2, 0) = "Range restriction in X (Observed score: ux)"
  Artifacts(4, 0) = "Reliability of X (Restricted population: rxx_i)"
  Artifacts(5, 0) = "Range restriction in X (True score: uT)"
  If CorrectRxx Then
    If RelUnrestx Then
      Artifacts(0, 1) = SumRxxFreq
      Artifacts(4, 1) = "Estimated"
    Else
      Artifacts(0, 1) = "Estimated"
      Artifacts(4, 1) = SumRxxFreq
    End If
    Artifacts(0, 2) = MeanRxxa
    Artifacts(0, 3) = SDRxxa
    Artifacts(0, 4) = MeanQualxa
    Artifacts(0, 5) = SDQualxa
    Artifacts(4, 2) = MeanRxxi
    Artifacts(4, 3) = SDRxxi
    Artifacts(4, 4) = MeanQualxi
    Artifacts(4, 5) = SDQualxi
  Else
    Artifacts(0, 1) = SumRxxFreq
    Artifacts(4, 1) = SumRxxFreq
    Artifacts(0, 2) = "--"
    Artifacts(0, 3) = "--"
    Artifacts(0, 4) = "--"
    Artifacts(0, 5) = "--"
    Artifacts(4, 2) = "--"
    Artifacts(5, 3) = "--"
    Artifacts(5, 4) = "--"
    Artifacts(5, 5) = "--"
  End If
  If CorrectRyy Then
    Artifacts(1, 1) = SumRyyFreq
    Artifacts(1, 2) = MeanRyy
    Artifacts(1, 3) = SDRyy
    Artifacts(1, 4) = MeanQualy
    Artifacts(1, 5) = SDQualy
  Else
    Artifacts(1, 1) = "SumRyyFreq"
    Artifacts(1, 2) = "--"
    Artifacts(1, 3) = "--"
    Artifacts(1, 4) = "--"
    Artifacts(1, 5) = "--"
  End If
  If Observedu Then
    Artifacts(2, 1) = SumUFreq
    Artifacts(5, 1) = "Estimated"
  Else
    Artifacts(2, 1) = "Estimated"
    Artifacts(5, 1) = SumUFreq
  End If
  Artifacts(2, 2) = Meanux
  Artifacts(2, 3) = SDux
  Artifacts(5, 2) = MeanuT
  Artifacts(5, 3) = SDuT
  
  ' Output and format artifacts table
  Worksheets("Output").Range("A12:F17") = Artifacts
  Worksheets("Output").Range("B12:B17").NumberFormat = "0"
  Worksheets("Output").Range("C12:F17").NumberFormat = ".00"
End If

' Prepare Supplemental Results Array
ReDim Supp(19, 8)
  'Extra intervals
  Supp(0, 0) = "Supplemental Results"
  Supp(1, 0) = "90% Conf. Int. (r)"
  Supp(2, 0) = "80% Cred. Int. (r)"
  Supp(3, 0) = "Lower CI (r)"
  Supp(4, 0) = "Upper CI (r)"
  Supp(5, 0) = "Lower CV (r)"
  Supp(6, 0) = "Upper CV (r)"

  Supp(3, 3) = "Lower CI (" & ChrW(961) & ")"
  Supp(4, 3) = "Upper CI (" & ChrW(961) & ")"
  Supp(5, 3) = "Lower CV (" & ChrW(961) & ")"
  Supp(6, 3) = "Upper CV (" & ChrW(961) & ")"

  Supp(3, 6) = "Lower CI (True validity)"
  Supp(4, 6) = "Upper CI (True validity)"
  Supp(5, 6) = "Lower CV (True validity)"
  Supp(6, 6) = "Upper CV (True validity)"

  Supp(1, 1) = Format(LoCImeanR, ".00") & ", " & Format(UpCImeanR, ".00")
  Supp(2, 1) = Format(LoCVmeanR, ".00") & ", " & Format(UpCVmeanR, ".00")
  Supp(3, 1) = LoCImeanR
  Supp(4, 1) = UpCImeanR
  Supp(5, 1) = LoCVmeanR
  Supp(6, 1) = UpCVmeanR

  Supp(3, 4) = LoCIrho
  Supp(4, 4) = UpCIrho
  Supp(5, 4) = LoCVrho
  Supp(6, 4) = UpCVrho

  Supp(3, 8) = LoCIrhoValidity
  Supp(4, 8) = UpCIrhoValidity
  Supp(5, 8) = LoCVrhoValidity
  Supp(6, 8) = UpCVrhoValidity

  ' Variance values
  Supp(8, 0) = "Observed variance of r"
  Supp(9, 0) = "Predicted variance of r"
  Supp(10, 0) = "Sampling error variance of r"
  Supp(11, 0) = "Variance due to artifact differences"
  Supp(12, 0) = "Residual variance of r"
  Supp(13, 0) = "True variance of " & ChrW(961)
  Supp(14, 0) = "Percent variance due to sampling error"
  Supp(15, 0) = "Percent variance accounted for"
  Supp(16, 0) = "Observed SD"
  Supp(17, 0) = "Predicted SD"

  Supp(8, 1) = ObsVar
  Supp(9, 1) = PredVar
  Supp(10, 1) = SampErrVar
  If Taylor Then
    Supp(11, 1) = "***"
    Supp(19, 0) = "*** Variance due to artifact differences is not separately estimated when using Taylor Series Approximation model"
  Else
    Supp(11, 1) = ArtVar
  End If
  Supp(12, 1) = ResVar
  Supp(13, 1) = VarRho
  Supp(14, 1) = PerVarSamp
  Supp(15, 1) = PerVarAcc
  Supp(16, 1) = SDobs
  Supp(17, 1) = SDpred
  
  ' Fail-safe values
  Supp(9, 3) = "Fail-safe k"
  Supp(10, 3) = "Fail-safe N"
  Supp(11, 3) = "Fail-safe threshold value"
  Supp(12, 3) = "Fail-safe mean file drawer value"
  
  Supp(9, 5) = kFS
  Supp(10, 5) = NFS
  Supp(11, 5) = rc
  Supp(12, 5) = rFS

' Output and format Supplemental Results
If Not CorrectRR Then
  Worksheets("Output").Range("A16:I35") = Supp
  Worksheets("Output").Range("A16").Font.Bold = True
  Worksheets("Output").Range("B19:I22,F24:F27,B32:B33").NumberFormat = ".00"
  Worksheets("Output").Range("B24:B29").NumberFormat = ".000"
  Worksheets("Output").Range("B30:B31").NumberFormat = "0%"
Else
  Worksheets("Output").Range("A19:I38") = Supp
  Worksheets("Output").Range("A19").Font.Bold = True
  Worksheets("Output").Range("B22:I25,F27:F30,B35:B36").NumberFormat = ".00"
  Worksheets("Output").Range("B27:B32").NumberFormat = ".000"
  Worksheets("Output").Range("B33:B34").NumberFormat = "0%"
End If
 
' ==========================
' ===== Print warnings =====
' ==========================

Worksheets("Alerts").Cells.ClearContents
Worksheets("Alerts").Cells.ClearFormats
If flags > 0 Then
  Worksheets("Alerts").Cells(1, 1).Value = "The following warnings were generated during the analyses:"
Else
  Worksheets("Alerts").Cells(1, 1).Value = "The analyses completed without errors or warnings."
End If
Worksheets("Alerts").Cells(1, 1).Font.Bold = True

i = 0
ReDim AlertFlags(4, 0)
If Not IsNull(FlagRxxOver) Then
  AlertFlags(i, 0) = FlagRxxOver
  i = i + 1
End If
If Not IsNull(FlagQualxUnder) Then
  AlertFlags(i, 0) = FlagQualxUnder
  i = i + 1
End If
If Not IsNull(FlagNoSDQualx) Then
  AlertFlags(i, 0) = FlagNoSDQualx
  i = i + 1
End If
If Not IsNull(FlagQualyUnder) Then
  AlertFlags(i, 0) = FlagQualyUnder
  i = i + 1
End If
If Not IsNull(FlagNoSDQualy) Then
  AlertFlags(i, 0) = FlagNoSDQualy
  i = i + 1
End If

Worksheets("Alerts").Range("A3:A7") = AlertFlags

' =========================
' ==== Display results ====
' =========================

Worksheets("Output").Select
Application.ScreenUpdating = True

End Sub