Attribute VB_Name = "test_z_os"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Function ts_z_os(data As Range, Optional mu = Null, Optional sigma = Null, Optional output = "all")

    If IsNull(mu) Then
        mu = (WorksheetFunction.Min(data) + WorksheetFunction.Max(data)) / 2
    End If
    
    If output = "mu" Then
        ts_z_os = mu
    Else
    
        'sample size (n), mean (avg) and standard deviation (s)
        n = WorksheetFunction.Count(data)
        avg = WorksheetFunction.Average(data)
        
        If IsNull(sigma) Then
            s = WorksheetFunction.StDev(data)
        Else
            s = sigma
        End If
        se = s / Sqr(n)
        
        If output = "se" Then
            ts_z_os = se
        Else
            'the t-value
            Z = (avg - mu) / se
            
            If output = "statistic" Then
                ts_z_os = Z
            
            Else
                'the p-value
                p = 2 * (1 - WorksheetFunction.NormSDist(Abs(Z)))
                        
                If output = "pvalue" Then
                    ts_z_os = p
            
                Else
                    'Results
                    Dim res(1 To 2, 1 To 5)
                    res(1, 1) = "mu"
                    res(1, 2) = "sample mean"
                    res(1, 3) = "statistic"
                    res(1, 4) = "p-value"
                    res(1, 5) = "test used"
                    res(2, 1) = mu
                    res(2, 2) = avg
                    res(2, 3) = Z
                    res(2, 4) = p
                    res(2, 5) = "one-sample z"
                    
                    ts_z_os = res
                End If
            End If
        End If
    End If

End Function
