Attribute VB_Name = "test_student_t_os"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076


Function ts_student_t_os(data As Range, Optional mu = "none", Optional output = "all")
Attribute ts_student_t_os.VB_Description = "perform a one-sample Student t test"
Attribute ts_student_t_os.VB_ProcData.VB_Invoke_Func = " \n14"
'perform a one-sample Student t test

    Dim n, df As Integer
    Dim avg, s, se, t, p As Double
    
    If mu = "none" Then
        mu = (WorksheetFunction.Min(data) + WorksheetFunction.Max(data)) / 2
    End If
    
    'sample size (n), mean (avg) and standard deviation (s)
    n = WorksheetFunction.Count(data)
    avg = WorksheetFunction.Average(data)
    s = WorksheetFunction.StDev(data)
    
    'degrees of freedom (df)
    df = n - 1
    
    If output = "df" Then
        ts_student_t_os = df
    Else
        se = s / Sqr(n)
        
        If output = "se" Then
            ts_student_t_os = se
        Else
            'the t-value
            t = (avg - mu) / se
            
            If output = "statistic" Then
                ts_student_t_os = t
            
            Else
                'the p-value
                p = WorksheetFunction.TDist(Abs(t), df, 2)
                        
                If output = "pvalue" Then
                    ts_student_t_os = p
            
                Else
                    'Results
                    Dim res(1 To 2, 1 To 6)
                    res(1, 1) = "mu"
                    res(1, 2) = "sample mean"
                    res(1, 3) = "statistic"
                    res(1, 4) = "df"
                    res(1, 5) = "p-value"
                    res(1, 6) = "test used"
                    res(2, 1) = mu
                    res(2, 2) = avg
                    res(2, 3) = t
                    res(2, 4) = df
                    res(2, 5) = p
                    res(2, 6) = "one-sample Student t"
                    
                    ts_student_t_os = res
                End If
            End If
        End If
    End If

End Function
