' Open Psychometric Meta-Analyis (d values)
' Created by Brenton M. Wiernik
' version 1.0.0

'    Open Psychometric Meta-Analysis (d values) -- VBA scripts for conducting psychometric
'    meta-analysis using Microsoft Excel.
'    Copyright (C) 2018 Brenton M. Wiernik.

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
' D = matrix of effect sizes and sample sizes
' RY = matrix of ryy values and frequencies

' Ntotal = total sample size (N)
' effectiveNtotal = effective total sample size (N) accounting for unequal groups
' sumD = weighted sumD of uncorrected d values
' meanD = weighted average uncorrected d value

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

' Delta = estimated mean true d effect size

' Datten = Attenuated d value for a particular ryy value
' DattenWeighted = Weighted attenuated d value for a particular ryy value
' SumDatten = Sum of attenuated d values
' SumDattenSq = Sum of squared attenuated d values
' ArtVar = Expected variance due to artifacts

' ResVar = Residual Variance of effect sizes
' SDres = Residual SD of effect sizes
' SDpred = Predicted SD of effect sizes
' PerVarAcc = Percent of variance in effect sizes accounted for

' SDdelta = True effect standard deviation

' SEmeanD = Standard error of mean d
' df = Degrees of freedom for confidence interval when k < 30
' Crit = Critical value for confidence interval when k < 30
' UpCIdelta = Upper value of 80% confidence interval for delta
' LoCIdelta = Lower value of 80% confidence interval for delta
' UpCImeanD = Upper value of 80% confidence interval for mean d
' LoCImeanD = Lower value of 80% confidence interval for mean d

' UpCVdelta = Upper value of 80% credibility interval for delta
' LoCVdelta = Lower value of 80% credibility interval for delta
' UpCVmeanD = Upper value of 80% credibility interval for mean d
' LoCVmeanD = Lower value of 80% credibility interval for mean d

' TODO: Add range restriction corrections
' TODO: Add indirect range restriction corrections via interactive method
' TODO: Compute sampling error based on unequal group sizes
' TODO: Allow weighting based on inverse variance

Sub MetaAnalysisD()

' Set up control parameters based on OS version and options set via worksheet controls
If Application.Version < 15 Then
  If Application.OperatingSystem Like "*Mac*" Then
    CompatMode = True
  ElseIf Application.Version < 14 Then
    CompatMode = True
  Else
    CompatMode = False
  End If
Else
  CompatMode = False
End If

If Worksheets("d - Artifact").Shapes("SepGroupN").ControlFormat.Value = 1 Then groupNs = True Else groupNs = False
If Worksheets("d - Artifact").Shapes("WtTotal").ControlFormat.Value = 1 Then
    Weights = "Total"
ElseIf Worksheets("d - Artifact").Shapes("WtUnit").ControlFormat.Value = 1 Then
    Weights = "Unit"
ElseIf Worksheets("d - Artifact").Shapes("WtInvSamp").ControlFormat.Value = 1 Then
    Weights = "InvSamp"
Else
    MsgBox ("Please select a weighting method.")
    Stop
End If
If Worksheets("ryy").Shapes("SpecDist").ControlFormat.Value = 1 Then SpecDist = True Else SpecDist = False

k = Application.Count(Worksheets("d - Artifact").Range("A:A"))
nRyy = Application.Count(Worksheets("ryy").Range("A:A"))

Ntotal = 0
effectiveNtotal = 0
Ngroup1 = 0
Ngroup2 = 0
If groupNs = True Then
    ReDim N(k, 3)
    For I = 1 To k
        N(I, 1) = Worksheets("d - Artifact").Cells(I + 1, 2).Value
        N(I, 2) = Worksheets("d - Artifact").Cells(I + 1, 3).Value
        N(I, 3) = (N(I, 1) * N(I, 2)) / (N(I, 1) + N(I, 2))
        Ntotal = Ntotal + N(I, 1) + N(I, 2)
        Ngroup1 = Ngroup1 + N(I, 1)
        Ngroup2 = Ngroup2 + N(I, 2)
        effectiveNtotal = effectiveNtotal + N(I, 3)
        Next I
Else
    ReDim N(k, 3)
    For I = 1 To k
        N(I, 1) = Worksheets("d - Artifact").Cells(I + 1, 2).Value
        N(I, 2) = 0
        N(I, 3) = N(I, 1)
        Ntotal = Ntotal + N(I, 1)
        Ngroup1 = ""
        Ngroup2 = ""
        effectiveNtotal = effectiveNtotal + N(I, 1)
        Next I
End If

ReDim D(k, 2)
For I = 1 To k
    D(I, 1) = Worksheets("d - Artifact").Cells(I + 1, 1).Value
    D(I, 2) = N(I, 3)
    Next I

If SpecDist = True Then
  ' TODO: Add support for specified artifact distributions
Else
  ReDim RY(nRyy, 2)
    For I = 1 To nRyy
    RY(I, 1) = Worksheets("ryy").Cells(I + 1, 1).Value
    RY(I, 2) = Worksheets("ryy").Cells(I + 1, 2).Value
    Next I
End If

' TODO: Correct d values for small sample size
        
' COMPUTE MEAN UNCORRECTED D
sumD = 0
For I = 1 To k
  sumD = sumD + D(I, 2) * D(I, 1)
  Next I
meanD = sumD / effectiveNtotal

' COMPUTE SAMPLING VAR OF OBS D'S
SampErrVar = (4 * (1 + (meanD ^ 2) / 8) * k) / effectiveNtotal
' TODO: Change variance to compute using unequal groups formula

' COMPUTE VAR OF OBS D'S
ObsSSQ = 0
For I = 1 To k
ObsSSQ = ObsSSQ + D(I, 2) * (D(I, 1) - meanD) ^ 2
Next I
ObsVar = ObsSSQ / effectiveNtotal
SDobs = Sqr(ObsVar)

' COMPUTE PERCENT VAR DUE TO SAMPLING ERROR
If ObsVar < 1E-16 Then
  PerVarSamp = "No Obs. Var."
Else
  PerVarSamp = (SampErrVar / ObsVar)
End If

' COMPUTE Ryy DISTRIBUTION
SumRyy = 0
SumRyyFreq = 0
SumRyySq = 0
SumQualy = 0
For I = 1 To nRyy
SumRyy = SumRyy + RY(I, 1) * RY(I, 2)
SumRyySq = SumRyySq + (RY(I, 1) ^ 2) * RY(1, 2)
SumQualy = SumQualy + Sqr(RY(I, 1)) * RY(I, 2)
SumRyyFreq = SumRyyFreq + RY(I, 2)
Next I
MeanRyy = SumRyy / SumRyyFreq
SDRyy = Sqr((SumRyySq / SumRyyFreq) - (SumRyy / SumRyyFreq) ^ 2)
MeanQualy = SumQualy / SumRyyFreq
SDQualy = Sqr((SumRyy / SumRyyFreq) - (SumQualy / SumRyyFreq) ^ 2)

' COMPUTE TRUE SCORE MEAN D
Delta = meanD / MeanQualy

' COMPUTE VAR DUE TO RYY DIFFS
Datten = 0
DattenWeighted = 0
SumDatten = 0
SumDattenSq = 0
For I = 1 To nRyy
Datten = Delta * Sqr(RY(I, 1))
DattenWeighted = Datten * RY(I, 2)
SumDatten = SumDatten + DattenWeighted
SumDattenSq = SumDattenSq + Datten ^ 2 * RY(I, 2)
Next I
ArtVar = (SumDattenSq / SumRyyFreq) - (SumDatten / SumRyyFreq) ^ 2

' COMPUTE RESIDUAL VAR and SD
ResVar = ObsVar - SampErrVar - ArtVar
If ResVar < 0 Then
  SDres = 0
Else
  SDres = Sqr(ResVar)
End If

' COMPUTE SD-PREDICTED
SDpred = Sqr(SampErrVar + ArtVar)

' COMPUTE PERCENT VAR ACC FOR
If ObsVar < 1E-16 Then
  PerVarAcc = "No Obs. Var."
Else
  PerVarAcc = ((SampErrVar + ArtVar) / ObsVar)
End If

' COMPUTE SD OF TRUE SCORE D'S
' TODO: Change this to the more robust method (based on .1*SDQualy steps up to 3*SDQualy) once range restriction is added
SDdelta = (Delta / meanD) * SDres

' Confidence Interval - 90% (meanD)
SEmeanD = SDobs / Sqr(k)
If k < 30 Then
  df = k - 1
  If CompatMode Then
    Crit = Application.TInv(0.1, df)
  Else
    Crit = Application.T_Inv_2T(0.1, df)
  End If
  UpCImeanD = meanD + Crit * SEmeanD
  LoCImeanD = meanD - Crit * SEmeanD
Else
  UpCImeanD = meanD + 1.64 * SEmeanD
  LoCImeanD = meanD - 1.64 * SEmeanD
End If

' Confidence Interval - 90%  (Delta)
' TODO: Change this to correct endpoints individually for unreliability and range restriction once RR is added
UpCIdelta = UpCImeanD * (Delta / meanD)
LoCIdelta = LoCImeanD * (Delta / meanD)

' Credibility Interval - 80% (meanD)
UpCVmeanD = meanD + 1.28 * SDres
LoCVmeanD = meanD - 1.28 * SDres

' Credibility Interval - 80% (Delta)
UpCVdelta = Delta + 1.28 * SDdelta
LoCVdelta = Delta - 1.28 * SDdelta

' TODO: Fix output page
    Worksheets("Output").Select
    Worksheets("Output").Cells.ClearContents

    Cells(1, 1).Value = "Meta-analysis Results"
    Cells(1, 1).Font.Bold = True
    Cells(2, 1).Value = "Standardized mean difference (Unbiased Cohen's d or Hedge's g)"
    Cells(3, 1).Value = "Corrected using artifact distribution"
    ' TODO: List the chosen checkbox options once they are implemented
    
    Cells(5, 1).Value = "Main Results"
    Cells(5, 1).Font.Bold = True
    Cells(6, 1).Value = "Recommended results table"
    Cells(6, 2).Value = "N (Total sample size)"
    Cells(7, 2).Value = Ntotal
    Cells(6, 3).Value = "Total Group 1 sample size"
    Cells(7, 3).Value = Ngroup1
    Cells(6, 4).Value = "Total Group 2 sample size"
    Cells(7, 4).Value = Ngroup2
    Cells(6, 5).Value = "Effective total sample size"
    Cells(7, 5).Value = effectiveNtotal
    Cells(6, 6).Value = "k (No. d values)"
    Cells(7, 6).Value = k
    Range("B7:F7").NumberFormat = "0"
    Cells(8, 2).Value = "Mean Uncorrected d"
    Cells(9, 2).Value = meanD
    Cells(8, 3).Value = "Observed SD of d"
    Cells(9, 3).Value = SDobs
    Cells(8, 4).Value = "Residual SD of d"
    Cells(9, 4).Value = SDres
    Cells(8, 5).Value = "Mean Corrected " & ChrW(948)
    Cells(9, 5).Value = Delta
    Cells(8, 6).Value = "SD of " & ChrW(948)
    Cells(9, 6).Value = SDdelta
    Cells(8, 7).Value = "90% Conf. Int. (" & ChrW(948) & ")"
    Cells(9, 7).Value = Application.Text(LoCIdelta, ".00") & ", " & Application.Text(UpCIdelta, ".00")
    Cells(8, 8).Value = "80% Cred. Int. (" & ChrW(948) & ")"
    Cells(9, 8).Value = Application.Text(LoCVdelta, ".00") & ", " & Application.Text(UpCVdelta, ".00")
    Range("B8:H8").NumberFormat = ".00"
    
    Cells(11, 1).Value = "Artifact Distribution"
    Cells(11, 1).Font.Bold = True
    Cells(12, 1).Value = "Recommended results table"
    Cells(12, 2).Value = "No. ryy values"
    Cells(12, 2).Value = SumRyyFreq
    Cells(12, 3).Value = "Mean ryy"
    Cells(13, 3).Value = MeanRyy
    Cells(12, 4).Value = "SD of ryy"
    Cells(13, 4).Value = SDRyy
    Cells(12, 5).Value = "Mean SQRT of ryy"
    Cells(13, 5).Value = MeanQualy
    Cells(12, 6).Value = "SD of SQRT of ryy"
    Cells(13, 6).Value = SDQualy
    Range("B13").NumberFormat = "0"
    Range("C13:F13").NumberFormat = ".00"
    
    Cells(15, 1).Value = "Supplemental Results"
    Cells(15, 1).Font.Bold = True
    Cells(16, 1).Value = "90% Conf. Int. (d)"
    Cells(16, 2).Value = Application.Text(LoCImeanD, ".00") & ", " & Application.Text(UpCImeanD, ".00")
    Cells(17, 1).Value = "80% Cred. Int. (d)"
    Cells(17, 2).Value = Application.Text(LoCVmeanD, ".00") & ", " & Application.Text(UpCVmeanD, ".00")
    Cells(18, 1).Value = "Lower CI (d)"
    Cells(18, 2).Value = LoCImeanD
    Cells(19, 1).Value = "Upper CI (d)"
    Cells(19, 2).Value = UpCImeanD
    Cells(20, 1).Value = "Lower CV (d)"
    Cells(20, 2).Value = LoCVmeanD
    Cells(21, 1).Value = "Upper CV (d)"
    Cells(21, 2).Value = UpCVmeanD
    Cells(23, 1).Value = "Lower CI (" & ChrW(948) & ")"
    Cells(23, 2).Value = LoCIdelta
    Cells(24, 1).Value = "Upper CI (" & ChrW(948) & ")"
    Cells(24, 2).Value = UpCIdelta
    Cells(25, 1).Value = "Lower CV (" & ChrW(948) & ")"
    Cells(25, 2).Value = LoCVdelta
    Cells(26, 1).Value = "Upper CV (" & ChrW(948) & ")"
    Cells(26, 2).Value = UpCVdelta
    Range("B16:B26").NumberFormat = ".00"
    
    Cells(28, 1).Value = "Observed variance of d"
    Cells(28, 2).Value = ObsVar
    Cells(29, 1).Value = "Sampling error variance of d"
    Cells(29, 2).Value = SampErrVar
    Cells(30, 1).Value = "Variance due to artifact differences"
    Cells(30, 2).Value = ArtVar
    Cells(31, 1).Value = "Residual variance"
    Cells(31, 2).Value = ResVar
    Cells(32, 1).Value = "Percent variance due to sampling error"
    Cells(32, 2).Value = PerVarSamp
    Cells(33, 1).Value = "Percent variance accounted for"
    Cells(33, 2).Value = PerVarAcc
    Range("B28:B31").NumberFormat = ".0000"
    Range("B32:B33").NumberFormat = "0%"
    
    Cells(35, 1).Value = "Observed SD"
    Cells(35, 2).Value = SDobs
    Cells(36, 1).Value = "Predicted SD"
    Cells(36, 2).Value = SDpred
    Range("B35:B36").NumberFormat = ".00"
End Sub


