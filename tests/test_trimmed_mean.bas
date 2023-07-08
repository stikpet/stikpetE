Attribute VB_Name = "test_trimmed_mean"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Function ts_trimmed_mean_os(data As Range, _
                            Optional mu = Null, _
                            Optional trimProp = 0.1, _
                            Optional se = "yuen", _
                            Optional output = "all")

    If IsNull(mu) Then
        mu = (WorksheetFunction.Min(data) + WorksheetFunction.Max(data)) / 2
    End If
    
    If output = "mu" Then
        ts_trimmed_mean_os = mu
    Else
    
        n = WorksheetFunction.Count(data)
        
        nt = n * trimProp / 2
        nl = WorksheetFunction.RoundDown(nt, 0)
        
        'determine the trimmed mean
        dataA = he_range_to_num_array(data)
        dataS = he_sort(dataA)
        s = 0
        For i = nl To n - nl - 1
            s = s + dataS(i)
        Next i
        mt = s / (n - 2 * nl)
        
        If output = "mt" Then
            ts_trimmed_mean_os = mt
        Else
        
            'number of scores after trimming
            nat = n - 2 * nl
            
            'determine the winsorized mean
            mw = (mt * nat + nl * (dataS(nl) + dataS(n - nl - 1))) / n
            
            'variance of winsorized data
            ss = 0
            For i = nl To n - nl - 1
                ss = ss + (dataS(i) - mw) ^ 2
            Next i
            ssdw = ss + nl * ((dataS(nl) - mw) ^ 2 + (dataS(n - nl - 1) - mw) ^ 2)
            varw = ssdw / (n - 1)
            
            If se = "yuen" Then
                se = (ssdw / (nat * (nat - 1))) ^ 0.5
            ElseIf se = "wilcox" Then
                se = (varw) ^ 0.5 / ((1 - trimProp) * (n ^ 0.5))
            End If
            
            If output = "se" Then
                ts_trimmed_mean_os = se
            Else
            
                tValue = (mt - mu) / se
                
                If output = "statistic" Then
                    ts_trimmed_mean_os = tValue
                
                Else
                
                df = nat - 1
                pValue = WorksheetFunction.TDist(Abs(tValue), df, 2)
                
                    If output = "pvalue" Then
                        s_trimmed_mean_os = pValue
                        
                    Else
                    
                        testUsed = "one-sample trimmed mean test"
                        
                        Dim results(1 To 2, 1 To 7)
                        results(1, 1) = "trim. mean"
                        results(1, 2) = "mu"
                        results(1, 3) = "SE"
                        results(1, 4) = "statistic"
                        results(1, 5) = "df"
                        results(1, 6) = "p-value"
                        results(1, 7) = "test used"
                        
                        results(2, 1) = mt
                        results(2, 2) = mu
                        results(2, 3) = se
                        results(2, 4) = tValue
                        results(2, 5) = df
                        results(2, 6) = pValue
                        results(2, 7) = testUsed
                        
                        ts_trimmed_mean_os = results
                    End If
                End If
            End If
        End If
    End If
    
End Function

