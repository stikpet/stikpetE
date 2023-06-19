Attribute VB_Name = "test_wilcoxon_os"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Public Sub ts_wilcoxon_os_addHelp()
Application.MacroOptions _
    Macro:="ts_wilcoxon_os", _
    Description:="one-sample Wilcoxon signed rank test", _
    category:=14, _
    ArgumentDescriptions:=Array( _
        "vertical specific range with data", _
        "optional vertical range with labels in order if data is non-numeric.", _
        "optional hypothesized median, otherwise the midrange will be used", _
        "optional boolean to use a tie correction (default is True)", _
        "optional method to use for approximation. Either " & Chr(34) & "wilcoxon" & Chr(34) & " (default), " & Chr(34) & "exact" & Chr(34) & ", " & Chr(34) & "imanz" & Chr(34) & " or " & Chr(34) & "imant" & Chr(34) & " for Iman's z or t approximation", _
        "optional method to deal with scores equal to mu. Either " & Chr(34) & "wilcoxon" & Chr(34) & " (default), " & Chr(34) & "pratt" & Chr(34) & " or " & Chr(34) & "zsplit" & Chr(34), _
        "optional boolean to use a continuity correction (default is False)", _
        "output to show, either " & Chr(34) & "all" & Chr(34) & "(default), " & Chr(34) & "pvalue" & Chr(34) & ", " & Chr(34) & "df" & Chr(34) & ", " & Chr(34) & "statistic" & Chr(34) & ", or " & Chr(34) & "w " & Chr(34))
                
End Sub

Function ts_wilcoxon_os(data As Range, _
                    Optional levels As Range, _
                    Optional mu = "none", _
                    Optional ties = True, _
                    Optional appr = "wilcoxon", _
                    Optional eqMed = "wilcoxon", _
                    Optional cc = False, _
                    Optional output = "all")


    'get data as numeric values
    If levels Is Nothing Then
        dataN = he_range_to_num_array(data)
    Else
        dataN = he_replace(data, levels)
    End If
    
    'sort the numeric values
    dataN = he_sort(dataN)
    
    
    If mu = "none" Then
        mu = (WorksheetFunction.Min(dataN) + WorksheetFunction.Max(dataN)) / 2
    End If
    
    Dim n, nr As Integer
    n = UBound(dataN, 1) - LBound(dataN, 1) + 1
    
    Dim i As Integer
    If eqMed = "wilcoxon" Or appr = "exact" Then
        nEqMed = 0
        For i = 0 To n - 1
            If dataN(i) = mu Then
                nEqMed = nEqMed + 1
            End If
        Next i
        
        nr = n - nEqMed
    Else
        nr = n
    End If
    
    Dim absDiffs() As Double
    ReDim absDiffs(0 To nr - 1)
    Dim scores() As Double
    ReDim scores(0 To nr - 1)
    Dim k As Integer
    k = 0
    i = 1
    Do While k < nr
        If eqMed = "wilcoxon" Then
            If dataN(i - 1) <> mu Then
                absDiffs(k) = Abs(dataN(i - 1) - mu)
                scores(k) = dataN(i - 1)
                k = k + 1
            End If
        
        Else
            absDiffs(k) = Abs(dataN(i - 1) - mu)
            scores(k) = dataN(i - 1)
            k = k + 1
        End If
        
        i = i + 1
    Loop
        
    
    'sort scores based on absolute differences
    changes = 1
    Do While changes <> 0
        changes = 0
        For i = 1 To nr - 1
            If absDiffs(i - 1) > absDiffs(i) Then
                ff1 = absDiffs(i)
                ff2 = scores(i)
                absDiffs(i) = absDiffs(i - 1)
                scores(i) = scores(i - 1)
                absDiffs(i - 1) = ff1
                scores(i - 1) = ff2
                
                changes = 1
            End If
        Next i
    Loop
    
    'we need the ranks for which we need the rank frequencies
    'store for each score how often it occurs
    'also check if ties actually occur
    maxRankFreq = 1
    Dim Rfreq As Variant
    ReDim Rfreq(1 To nr, 1 To 3)
    For i = 0 To nr - 1
        For j = 0 To nr - 1
            If absDiffs(j) = absDiffs(i) Then
                freq = freq + 1
            End If
        Next j
        
        Rfreq(i + 1, 1) = absDiffs(i)
        Rfreq(i + 1, 2) = freq
        
        If freq > maxRankFreq Then
            maxRankFreq = freq
        End If
        
        freq = 0
    Next i
    
    'now for the ranks and sum of ranks
    nD0 = 0
    Rsum = 0
    Wmin = 0
    Rd0 = 0 'for sum of ranks of differences of 0
    ReDim r(1 To nr, 1 To 3)
    r(1, 1) = Rfreq(1, 1)
    r(1, 2) = Rfreq(1, 2)
    
    If Rfreq(1, 2) = 1 Then
        r(1, 3) = 1
        Else
        r(1, 3) = (1 + 1 + Rfreq(1, 2) - 1) / 2
    End If
    
    If scores(0) > mu Then
        Rsum = Rsum + r(1, 3)
    ElseIf scores(0) = mu Then
        nD0 = nD0 + 1
        Rd0 = Rd0 + r(1, 3)
    Else
        Wmin = Wmin + r(1, 3)
    End If
    
    
    For i = 2 To nr
        r(i, 1) = Rfreq(i, 1)
        r(i, 2) = Rfreq(i, 2)
        
        If Rfreq(i, 2) = 1 Then
            r(i, 3) = i
        ElseIf Rfreq(i, 1) <> Rfreq(i - 1, 1) Then
            r(i, 3) = (i + i + Rfreq(i, 2) - 1) / 2
        Else
            r(i, 3) = r(i - 1, 3)
        End If
            
        If scores(i - 1) > mu Then
            Rsum = Rsum + r(i, 3)
        ElseIf scores(i - 1) = mu Then
            nD0 = nD0 + 1
            Rd0 = Rd0 + r(i, 3)
        Else
            Wmin = Wmin + r(i, 3)
        End If
        
    Next i
    
    If eqMed = "wilcoxon" Or eqMed = "pratt" Or appr = "exact" Then
        w = Rsum
    ElseIf eqMed = "zsplit" Then
        testUsed = testUsed + ", z-split method for equal to hyp. med."
        w = Rd0 / 2 + Rsum
    End If
    
    If eqMed = "zsplit" Then
        nr = n
    End If
    
    f1 = nr + 1
    s2 = nr * f1 * (2 * nr + 1) / 24
    rAvg = nr * f1 / 4
    
    
    If eqMed = "pratt" Then
        testUsed = testUsed + ", Pratt method for equal to hyp. med. (inc. Cureton adjustment for normal approximation)"
        'normal approximation adjustment based on Cureton (1967)
        s2 = s2 - nD0 * (nD0 + 1) * (2 * nD0 + 1) / 24
        rAvg = (nr * f1 - nD0 * (nD0 + 1)) / 4
    End If
    
    If ties = True Then
        testUsed = testUsed + ", ties correction applied"
        'ties correction
        t = 0
        For i = 1 To nr - 1
            If Rfreq(i, 1) <> Rfreq(i + 1, 1) Then
                If eqMed = "wilcoxon" Or eqMed = "pratt" Then
                'exclude those equal to hypothesized median
                    If Rfreq(i, 1) <> 0 Then
                        t = t + (Rfreq(i, 2) ^ 3 - Rfreq(i, 2)) / 48
                    End If
                Else
                    t = t + (Rfreq(i, 2) ^ 3 - Rfreq(i, 2)) / 48
                End If
                
            End If
        Next i
        If Rfreq(nr - 1, 1) = Rfreq(nr, 1) Then
            t = t + (Rfreq(i, 2) ^ 3 - Rfreq(i, 2)) / 48
        End If
        
        s2 = s2 - t
    End If
    
    If appr = "exact" Then
            If maxRankFreq > 1 Then
                testUsed = "ties occur, cannot compute exact method"
                pVal = "n.a."
                df = "n.a."
                statistic = "n.a."
            Else
                statistic = WorksheetFunction.Min(w, Wmin)
                pVal = di_wcdf(statistic, nr) * 2
                testUsed = "one-sample Wilcoxon signed rank exact test"
                df = "n.a."
            End If
    Else
        se = Sqr(s2)
        
        If cc = True Then
            num = Abs(w - rAvg) - 0.5
        Else
            num = Abs(w - rAvg)
        End If
            
        If appr = "imant" Then
            testUsed = testUsed + ", using Iman's t approximation"
            tValue = num / Sqr((s2 * nr - (w - rAvg) ^ 2) / (nr - 1))
            df = nr - 1
            statistic = tValue
        Else:
            zValue = num / se
            statistic = zValue
            df = "n.a."
        End If
        
        If appr = "imanz" Then
            testUsed = testUsed + ", using Iman's z approximation"
            zValue = zValue / 2 * (1 + Sqr((nr - 1) / (nr - zValue ^ 2)))
            statistic = zValue
        End If
        
        If appr = "imant" Then
            pVal = WorksheetFunction.T_Dist_2T(Abs(tValue), df)
        Else
            pVal = 2 * (1 - WorksheetFunction.Norm_S_Dist(Abs(zValue), True))
        End If
    End If

    If output = "w" Then
        ts_wilcoxon_os = w
    ElseIf output = "statistic" Then
        ts_wilcoxon_os = statistic
    ElseIf output = "pvalue" Then
        ts_wilcoxon_os = pVal
    ElseIf output = "df" Then
        ts_wilcoxon_os = df
    Else
        Dim res(1 To 2, 1 To 5)
        res(1, 1) = "W"
        res(1, 2) = "statistic"
        res(1, 3) = "df"
        res(1, 4) = "p-value"
        res(1, 5) = "test"
        res(2, 1) = w
        res(2, 2) = statistic
        res(2, 3) = df
        res(2, 4) = pVal
        res(2, 5) = testUsed
        
        ts_wilcoxon_os = res
    End If

End Function

