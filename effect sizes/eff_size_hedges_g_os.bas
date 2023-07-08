Attribute VB_Name = "eff_size_hedges_g_os"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Function es_hedges_g_os(data As Range, Optional mu = Null, Optional appr = Null, Optional output = "all")
' Calculates Hedges G for a one-sample
' appr can be set to 'exact', 'hedges', 'durlak', or 'xue'.

    If IsNull(mu) Then
        mu = (WorksheetFunction.Min(data) + WorksheetFunction.Max(data)) / 2
    End If
    
    'Calculate mean, difference with hyp. mean and standard deviation
    m = WorksheetFunction.Average(data)
    dif = m - mu
    s = WorksheetFunction.StDev(data)
    
    'Determine Cohen's d
    d = dif / s
    
    n = WorksheetFunction.Count(data)
    df = n - 1
    
    m = df / 2
    
    If IsNull(appr) And m <= 171 Then
        g = d * WorksheetFunction.Gamma(m) / (WorksheetFunction.Gamma(m - 0.5) * m ^ 0.5)
        Comment = "exact"
    ElseIf appr = "hedges" Or (m > 171 And appr = "auto") Then
        g = d * (1 - 3 / (4 * df - 1))
        Comment = "Hedges approximation"
    ElseIf appr = "durlak" Then
        g = d * (n - 3) / (n - 2.25) * ((n - 2) / n) ^ 0.5
        Comment = "Durlak approximation"
    Else
        g = d * (1 - 9 / df + 69 / (2 * df ^ 2) - 72 / (df ^ 3) + 687 / (8 * df ^ 4) - 441 / (8 * df ^ 5) + 247 / (16 * df ^ 6)) ^ (1 / 12)
        Comment = "Xue approximation"
    End If
        
    'Results
    If out = "value" Then
        es_hedges_g_os = g
    Else
        Dim res(1 To 2, 1 To 3)
        res(1, 1) = "mu"
        res(1, 2) = "g"
        res(1, 3) = "version"
        res(2, 1) = mu
        res(2, 2) = g
        res(2, 3) = Comment
        es_hedges_g_os = res
    End If

End Function
