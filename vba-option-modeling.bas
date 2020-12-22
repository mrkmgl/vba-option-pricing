Attribute VB_Name = "Module1"
Option Explicit

Sub OptionPricing_CRR_BS()

Dim OptionType As Integer
Dim S0, K, r, v, t, u, d, DeltaT As Double
Dim D1, D2 As Double
Dim N, i, b, x, j As Long
Dim p, PayoffEUCall(0 To 999), PayoffEUPut As Variant
Dim PayoffUSCall(0 To 999), PayoffUSPut(0 To 999) As Variant
Dim PayoffTempCall(0 To 999), PayoffTempPut(0 To 999) As Variant
Dim BSCall, BSPut As Double

'The user can define all the variables, including the type of model to use. S/he has just to choose between 1, 6. This avoids creating If statements.
OptionType = InputBox("Press a number listed below" & vbCrLf & "1. European Call" & vbCrLf & "2. European Put" & vbCrLf & "3. American Call" & vbCrLf & "4. American Put" & vbCrLf & "5. Black Scholes Call" & vbCrLf & "6. Black Scholes Put")
S0 = InputBox("What is the Initial Stock Price?")
K = InputBox("What is the Strike Price?")
r = InputBox("Define the Interest Rate (in %)") / 100 'making it relative
v = InputBox("Define the Stock Volatility (in %)") / 100 'making it relative
t = InputBox("What is the option Maturity (base 365)?") / 365 'making it in 365 base

'We start with the Black-Scholes to make the code more clear.
'Black-Scholes formula application
With Application
    D1 = (.Ln(S0 / K) + (r + (v ^ 2) / 2) * t) / (v * Sqr(t))
    D2 = (.Ln(S0 / K) + (r - (v ^ 2) / 2) * t) / (v * Sqr(t))
End With

'If BS-Call is called, we apply the BS formula for the call
If OptionType = 5 Then
    With Application
        BSCall = S0 * .Norm_S_Dist(D1, True) - K * Exp(-r * t) * .Norm_S_Dist(D2, True)
    End With
    MsgBox "The Black-Scholes Call Price is " & BSCall
    Exit Sub
End If

'If BS-Put is called, we apply the BS formula for the put
If OptionType = 6 Then
    With Application
        BSPut = K * Exp(-r * t) * .Norm_S_Dist(-D2, True) - S0 * .Norm_S_Dist(-D1, True)
    End With
    MsgBox "The Black-Scholes Put Price is " & BSPut
    Exit Sub
End If

'We continue with the CRR model, therefore we ask for the number of steps not required by the BS formula
'All the price values are stored independently whether the options is called or not.
N = InputBox("Define the number of steps")
DeltaT = t / N 'time variation
u = Exp(v * Sqr(DeltaT)) 'up movement
d = Exp(-v * Sqr(DeltaT)) 'down movement
p = ((Exp(r * DeltaT)) - d) / (u - d) 'risk neutral probability

'The first loop goes from 1 to N+1 (or 0 to N).
'It calculates all the payoff for each state.
'Therefore, we get the expectation following the theoretical approach
For i = 1 To N + 1
    b = N + 1 - i
    PayoffEUCall(i) = (S0 * WorksheetFunction.Power(u, b) * WorksheetFunction.Power(d, i - 1)) - K
    If PayoffEUCall(i) <= 0 Then PayoffEUCall(i) = 0
    
    'Payoffs are the same, we are not deciding to sell or continue the investment in this loop.
    PayoffUSCall(i) = PayoffEUCall(i)
    
    PayoffUSPut(i) = K - (S0 * WorksheetFunction.Power(u, b) * WorksheetFunction.Power(d, i - 1))
    If PayoffUSPut(i) <= 0 Then PayoffUSPut(i) = 0
Next i

'This loop goes back from N to 1.
'It allocates the payoff and make decisions based on the the payoff and strike value
'Therefore, the values are stored in the Payoff Tables using the theoretical approach
For j = N To 1 Step -1
    For x = 1 To j
        PayoffEUCall(x) = Exp(-r * DeltaT) * ((p * PayoffEUCall(x)) + ((1 - p) * PayoffEUCall(x + 1)))
        
        'For the American case we make decisions using the PayoffTempCall/Put Table.
        'If the difference between the calculated payoff and that table is less or equal than 0 ==> we choose the latter.
        PayoffUSCall(x) = Exp(-r * DeltaT) * ((p * PayoffUSCall(x)) + ((1 - p) * PayoffUSCall(x + 1)))
        PayoffTempCall(x) = (S0 * WorksheetFunction.Power(u, j - x + 1) * WorksheetFunction.Power(d, x - 1)) - K
        If PayoffUSCall(x) - PayoffTempCall(x) <= 0 Then PayoffUSCall(x) = PayoffTempCall(x)
        
        PayoffUSPut(x) = Exp(-r * DeltaT) * ((p * PayoffUSPut(x)) + ((1 - p) * PayoffUSPut(x + 1)))
        PayoffTempPut(x) = K - (S0 * WorksheetFunction.Power(u, j - x + 1) * WorksheetFunction.Power(d, x - 1))
        If PayoffUSPut(x) - PayoffTempPut(x) <= 0 Then PayoffUSPut(x) = PayoffTempPut(x)
    Next x
Next j

PayoffEUPut = PayoffEUCall(1) - S0 + (K * Exp(-r * t))

If OptionType = 1 Then MsgBox "The European Call Price is " & PayoffEUCall(1)
If OptionType = 2 Then MsgBox "The European Put Price is " & PayoffEUPut
If OptionType = 3 Then MsgBox "The American Call price is " & PayoffUSCall(1)
If OptionType = 4 Then MsgBox "The American Put price is " & PayoffUSPut(1)

End Sub

Sub Sensitivity_Analysis()

Dim S0, K, r, v, t, u, d, D1, D2, DeltaT As Double
Dim N, i, b, x, y, Row As Long
Dim p, PayoffEUCall(0 To 999) As Variant
Dim co1, co2, co3, co4, co5, co6 As ChartObject
Dim ChartS0, ChartK, ChartR, ChartVol, ChartMat, ChartSteps As Object
Dim sc1, sc2, sc3, sc4, sc5, sc6 As SeriesCollection
Dim ser1, ser2, ser3, ser4, ser5, ser6, ser7 As Series

'We initiate letting variate the initial stock price (S0)
Cells(1, 1).Value = "Initial Stock Price" 'rename the column
Cells(1, 2).Value = "EU Call Price CRR" 'rename the column

'This loop will do the calculation for each row (100 data-points)
For Row = 1 To 100
K = 50
r = 0.05
t = 1
N = 10
v = 0.2
S0 = 30 + ((40 / 99) * (Row - 1)) 'the variation range has been divided by the number of data-points (uniformly distributed variation)

'We apply the same formula developed in the previous macro.
'For reference, please check above.
DeltaT = t / N
u = Exp(v * Sqr(DeltaT))
d = Exp(-v * Sqr(DeltaT))
p = ((Exp(r * DeltaT)) - d) / (u - d)
For i = 1 To N + 1
    b = N + 1 - i
    PayoffEUCall(i) = (S0 * WorksheetFunction.Power(u, b) * WorksheetFunction.Power(d, i - 1)) - K
    If PayoffEUCall(i) <= 0 Then PayoffEUCall(i) = 0
Next i
For y = N To 1 Step -1
    For x = 1 To y
        PayoffEUCall(x) = Exp(-r * DeltaT) * ((p * PayoffEUCall(x)) + ((1 - p) * PayoffEUCall(x + 1)))
    Next x
Next y
Cells(Row + 1, 1).Value = S0 'we store the value
Cells(Row + 1, 2).Value = PayoffEUCall(1) 'we store the value
Next Row

'We create the chart to show the sensitivity
Set co1 = ActiveSheet.ChartObjects.Add(Range("A103").Left, Range("A103").Top, 400, 200) 'set where the chart will be located in the Sheet
co1.Name = "Initial Stock Price Sensitivity"

Set ChartS0 = co1.Chart
With ChartS0

    .HasLegend = False
    .HasTitle = True
    .ChartTitle.Text = "Initial Stock Price Sensitivity"
    
    Set sc1 = .SeriesCollection
    Set ser1 = sc1.NewSeries
    
    With ser1
    
        .XValues = Range(Range("A1").Offset(1, 0), Range("A1").End(xlDown))
        .Values = Range(Range("A1").Offset(1, 1), Range("A1").Offset(1, 1).End(xlDown))
        .ChartType = xlLine
        
    End With

End With

'We continue letting variate the strike price (K)
Cells(1, 3).Value = "Strike Price" 'renaming the column
Cells(1, 4).Value = "EU Call Price CRR" 'renaming the column

For Row = 1 To 100
K = 30 + ((40 / 99) * (Row - 1)) 'the variation range has been divided by the number of data-points (uniformly distributed variation)
r = 0.05
t = 1
N = 10
v = 0.2
S0 = 50

'We apply the same formula developed in the previous macro.
'For reference, please check above.
DeltaT = t / N
u = Exp(v * Sqr(DeltaT))
d = Exp(-v * Sqr(DeltaT))
p = ((Exp(r * DeltaT)) - d) / (u - d)
For i = 1 To N + 1
    b = N + 1 - i
    PayoffEUCall(i) = (S0 * WorksheetFunction.Power(u, b) * WorksheetFunction.Power(d, i - 1)) - K
    If PayoffEUCall(i) <= 0 Then PayoffEUCall(i) = 0
Next i
For y = N To 1 Step -1
    For x = 1 To y
        PayoffEUCall(x) = Exp(-r * DeltaT) * ((p * PayoffEUCall(x)) + ((1 - p) * PayoffEUCall(x + 1)))
    Next x
Next y
Cells(Row + 1, 3).Value = K 'we store the value along the column
Cells(Row + 1, 4).Value = PayoffEUCall(1) 'we store the value along the column
Next Row

'We create the chart to show the sensitivity
Set co2 = ActiveSheet.ChartObjects.Add(Range("G103").Left, Range("G103").Top, 400, 200) 'set where the chart will be located in the Sheet
co2.Name = "Strike Price Sensitivity"

Set ChartK = co2.Chart
With ChartK
    
    .HasLegend = False
    .HasTitle = True
    .ChartTitle.Text = "Strike Price Sensitivity"
    
    Set sc2 = .SeriesCollection
    Set ser2 = sc2.NewSeries
    
    With ser2
    
        .XValues = Range(Range("C1").Offset(1, 0), Range("C1").End(xlDown))
        .Values = Range(Range("C1").Offset(1, 1), Range("C1").Offset(1, 1).End(xlDown))
        .ChartType = xlLine
        
    End With

End With

'We continue letting variate the interest rate
Cells(1, 5).Value = "Interest Rate" 'renaming the column
Cells(1, 6).Value = "EU Call Price CRR" 'renaming the column

For Row = 1 To 100
K = 50
r = 0.03 + ((0.04 / 99) * (Row - 1)) 'the variation range has been divided by the number of data-points (uniformly distributed variation)
t = 1
N = 10
v = 0.2
S0 = 50

'We apply the same formula developed in the previous macro.
'For reference, please check above.
DeltaT = t / N
u = Exp(v * Sqr(DeltaT))
d = Exp(-v * Sqr(DeltaT))
p = ((Exp(r * DeltaT)) - d) / (u - d)
For i = 1 To N + 1
    b = N + 1 - i
    PayoffEUCall(i) = (S0 * WorksheetFunction.Power(u, b) * WorksheetFunction.Power(d, i - 1)) - K
    If PayoffEUCall(i) <= 0 Then PayoffEUCall(i) = 0
Next i
For y = N To 1 Step -1
    For x = 1 To y
        PayoffEUCall(x) = Exp(-r * DeltaT) * ((p * PayoffEUCall(x)) + ((1 - p) * PayoffEUCall(x + 1)))
    Next x
Next y
Cells(Row + 1, 5).Value = r 'we store the value along the column
Cells(Row + 1, 6).Value = PayoffEUCall(1) 'we store the value along the column
Next Row

'We create the chart to show the sensitivity
Set co3 = ActiveSheet.ChartObjects.Add(Range("M103").Left, Range("M103").Top, 400, 200) 'set where the chart will be located in the Sheet
co3.Name = "Interest Rate Sensitivity"

Set ChartR = co3.Chart
With ChartR
    
    .HasLegend = False
    .HasTitle = True
    .ChartTitle.Text = "Interest Rate Sensitivity"
    
    Set sc3 = .SeriesCollection
    Set ser3 = sc3.NewSeries
    
    With ser3
    
        .XValues = Range(Range("E1").Offset(1, 0), Range("E1").End(xlDown))
        .Values = Range(Range("E1").Offset(1, 1), Range("E1").Offset(1, 1).End(xlDown))
        .ChartType = xlLine
        
    End With

End With

'We continue letting variate the stock volatility
Cells(1, 7).Value = "Stock Volatility" 'renaming the column
Cells(1, 8).Value = "EU Call Price CRR" 'renaming the column

For Row = 1 To 100
K = 50
r = 0.05
t = 1
N = 10
v = 0.05 + ((0.3 / 99) * (Row - 1)) 'the variation range has been divided by the number of data-points (uniformly distributed variation)
S0 = 50

'We apply the same formula developed in the previous macro.
'For reference, please check above.
DeltaT = t / N
u = Exp(v * Sqr(DeltaT))
d = Exp(-v * Sqr(DeltaT))
p = ((Exp(r * DeltaT)) - d) / (u - d)
For i = 1 To N + 1
    b = N + 1 - i
    PayoffEUCall(i) = (S0 * WorksheetFunction.Power(u, b) * WorksheetFunction.Power(d, i - 1)) - K
    If PayoffEUCall(i) <= 0 Then PayoffEUCall(i) = 0
Next i
For y = N To 1 Step -1
    For x = 1 To y
        PayoffEUCall(x) = Exp(-r * DeltaT) * ((p * PayoffEUCall(x)) + ((1 - p) * PayoffEUCall(x + 1)))
    Next x
Next y
Cells(Row + 1, 7).Value = v 'we store the value along the column
Cells(Row + 1, 8).Value = PayoffEUCall(1) 'we store the value along the column
Next Row

'We create the chart to show the sensitivity
Set co4 = ActiveSheet.ChartObjects.Add(Range("U103").Left, Range("U103").Top, 400, 200) 'set where the chart will be located in the Sheet
co4.Name = "Stock Volatility Sensitivity"

Set ChartVol = co4.Chart
With ChartVol
    
    .HasLegend = False
    .HasTitle = True
    .ChartTitle.Text = "Stock Volatility Sensitivity"
    
    Set sc4 = .SeriesCollection
    Set ser4 = sc4.NewSeries
    
    With ser4
    
        .XValues = Range(Range("G1").Offset(1, 0), Range("G1").End(xlDown))
        .Values = Range(Range("G1").Offset(1, 1), Range("G1").Offset(1, 1).End(xlDown))
        .ChartType = xlLine
        
    End With

End With

'We continue letting variate the maturity
Cells(1, 9).Value = "Maturity" 'renaming the column
Cells(1, 10).Value = "EU Call Price CRR" 'renaming the column

For Row = 1 To 100
K = 50
r = 0.05
t = 0.1 + ((1.8 / 99) * (Row - 1)) 'the variation range has been divided by the number of data-points (uniformly distributed variation)
N = 10
v = 0.2
S0 = 50

'We apply the same formula developed in the previous macro.
'For reference, please check above.
DeltaT = t / N
u = Exp(v * Sqr(DeltaT))
d = Exp(-v * Sqr(DeltaT))
p = ((Exp(r * DeltaT)) - d) / (u - d)
For i = 1 To N + 1
    b = N + 1 - i
    PayoffEUCall(i) = (S0 * WorksheetFunction.Power(u, b) * WorksheetFunction.Power(d, i - 1)) - K
    If PayoffEUCall(i) <= 0 Then PayoffEUCall(i) = 0
Next i
For y = N To 1 Step -1
    For x = 1 To y
        PayoffEUCall(x) = Exp(-r * DeltaT) * ((p * PayoffEUCall(x)) + ((1 - p) * PayoffEUCall(x + 1)))
    Next x
Next y
Cells(Row + 1, 9).Value = t 'we store the value along the column
Cells(Row + 1, 10).Value = PayoffEUCall(1) 'we store the value along the column
Next Row

'We create the chart to show the sensitivity
Set co5 = ActiveSheet.ChartObjects.Add(Range("AC103").Left, Range("AC103").Top, 400, 200) 'set where the chart will be located in the Sheet
co5.Name = "Maturity Sensitivity"

Set ChartMat = co5.Chart
With ChartMat
    
    .HasLegend = False
    .HasTitle = True
    .ChartTitle.Text = "Maturity Sensitivity"
    
    Set sc5 = .SeriesCollection
    Set ser5 = sc5.NewSeries
    
    With ser5
    
        .XValues = Range(Range("I1").Offset(1, 0), Range("I1").End(xlDown))
        .Values = Range(Range("I1").Offset(1, 1), Range("I1").Offset(1, 1).End(xlDown))
        .ChartType = xlLine
        
    End With

End With

'We finish letting variate the number of steps.
'We also compare the results with the Black-Scholes value for graphical purposes
Cells(1, 11).Value = "Number of Steps" 'renaming the column
Cells(1, 12).Value = "EU Call Price CRR" 'renaming the column
Cells(1, 13).Value = "EU Call Price BS" 'renaming the column

For Row = 1 To 100
K = 50
r = 0.05
t = 1
N = 1 + ((99 / 99) * (Row - 1)) 'the variation range has been divided by the number of data-points (uniformly distributed variation)
v = 0.2
S0 = 50

'We apply the same formula developed in the previous macro.
'For reference, please check above.
DeltaT = t / N
u = Exp(v * Sqr(DeltaT))
d = Exp(-v * Sqr(DeltaT))
p = ((Exp(r * DeltaT)) - d) / (u - d)
For i = 1 To N + 1
    b = N + 1 - i
    PayoffEUCall(i) = (S0 * WorksheetFunction.Power(u, b) * WorksheetFunction.Power(d, i - 1)) - K
    If PayoffEUCall(i) <= 0 Then PayoffEUCall(i) = 0
Next i
For y = N To 1 Step -1
    For x = 1 To y
        PayoffEUCall(x) = Exp(-r * DeltaT) * ((p * PayoffEUCall(x)) + ((1 - p) * PayoffEUCall(x + 1)))
    Next x
Next y
Cells(Row + 1, 11).Value = N 'we store the value along the column
Cells(Row + 1, 12).Value = PayoffEUCall(1) 'we store the value along the column

'We apply the same formula developed in the previous macro (Black-Scholes).
'For reference, please check above.
With Application
    D1 = (.Ln(S0 / K) + (r + (v ^ 2) / 2) * t) / (v * Sqr(t))
    D2 = (.Ln(S0 / K) + (r - (v ^ 2) / 2) * t) / (v * Sqr(t))
    Cells(Row + 1, 13).Value = S0 * .Norm_S_Dist(D1, True) - K * Exp(-r * t) * .Norm_S_Dist(D2, True) 'we store the value along the column
End With

Next Row

'We create the chart to show the sensitivity
Set co6 = ActiveSheet.ChartObjects.Add(Range("A118").Left, Range("AC118").Top, 600, 200) 'set where the chart will be located in the Sheet
co6.Name = "Steps Sensitivity & Comparison BS"

Set ChartSteps = co6.Chart
With ChartSteps
    
    .HasLegend = True
    .HasTitle = True
    .ChartTitle.Text = "Steps Sensitivity & Comparison BS"
    .Axes(xlValue).MinimumScale = 4.7
    .Axes(xlValue).MaximumScale = 6.3
    
    Set sc6 = .SeriesCollection
    Set ser6 = sc6.NewSeries
    
    With ser6
    
        .XValues = Range(Range("K1").Offset(1, 0), Range("K1").End(xlDown))
        .Values = Range(Range("K1").Offset(1, 1), Range("K1").Offset(1, 1).End(xlDown))
        .Name = "Price CRR"
        .ChartType = xlLine
        
    End With
    
    Set ser7 = sc6.NewSeries
    
    With ser7
            
        .XValues = Range(Range("K1").Offset(1, 0), Range("K1").End(xlDown))
        .Values = Range(Range("L1").Offset(1, 1), Range("L1").Offset(1, 1).End(xlDown))
        .Name = "Price BS"
        .ChartType = xlLine
    
    End With
    
End With

End Sub
