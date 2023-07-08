Attribute VB_Name = "eff_size_cohen_d_os"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076


Function es_cohen_d_os(data As Range, Optional mu = Null)
Attribute es_cohen_d_os.VB_Description = "Calculate Cohen d'"
Attribute es_cohen_d_os.VB_ProcData.VB_Invoke_Func = " \n14"
'Function to calculate Cohen's d for a one-sample t-test
'Parameters: data as a list of numbers and hypMean as the hypothesized mean

    Dim m, dif, s, d As Double
    
    If IsNull(mu) Then
        mu = (WorksheetFunction.Min(data) + WorksheetFunction.Max(data)) / 2
    End If
    
    'Calculate mean, difference with hyp. mean and standard deviation
    m = WorksheetFunction.Average(data)
    dif = m - mu
    s = WorksheetFunction.StDev(data)
    
    'Determine Cohen's d
    dos = Abs(dif) / s
    
    'Results
    es_cohen_d_os = dos


End Function
