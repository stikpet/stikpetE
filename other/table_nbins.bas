Attribute VB_Name = "table_nbins"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

'Function is used in the tab_frequency_bins() function

Function tab_nbins(data As Range, _
                Optional method = "src", _
                Optional maxBins = 100, _
                Optional adjust = 1, _
                Optional qmethod = "cdf")

    n = WorksheetFunction.Count(data)
    
    If method = "src" Then
        k = Sqr(n)
    
    ElseIf method = "sturges" Then
        k = Log(n) / Log(2) + 1
    
    ElseIf method = "qr" Then
        k = 2.5 * n ^ (1 / 4)
    
    ElseIf method = "rice" Then
        k = 2 * (n ^ (1 / 3))

    'Terrell and Scott
    ElseIf method = "ts" Then
        k = (2 * n) ^ (1 / 3)

    'Exponential
    ElseIf method = "exp" Then
        k = Log(n) / Log(2)

    'Exponential
    ElseIf method = "velleman" Then
        If n <= 100 Then
            k = 2 * Sqr(n)
        Else
            k = 10 * Log(n) / Log(10)
        End If
            
    'Doane
    ElseIf method = "doane" Then
        avg = WorksheetFunction.Sum(data) / n
        s = 0
        s2 = 0
        For Each Cell In data
            If Not IsEmpty(Cell.value) And WorksheetFunction.IsNumber(Cell.value) Then
                s = s + (Cell.value - avg) ^ 2
                s2 = s2 + (Cell.value - avg) ^ 3
            End If
        Next Cell
        sigSkew = (6 * (n - 2) / ((n + 1) * (n + 3))) ^ 0.5
        sPop = (s / n) ^ 0.5
        g1 = s2 / ((n) * sPop ^ 3)
        k = 1 + Log(n) / Log(2) + Log(Abs(g1) / sigSkew) / Log(2)
        
    Else
        r = WorksheetFunction.Max(data) - WorksheetFunction.Min(data)

        'Scott
        If method = "scott" Then
            avg = WorksheetFunction.Sum(data) / n
            s = 0
            For Each Cell In data
                If Not IsEmpty(Cell.value) And WorksheetFunction.IsNumber(Cell.value) Then
                    s = s + (Cell.value - avg) ^ 2
                End If
            Next Cell
            sd = (s / (n - 1)) ^ 0.5
            h = 3.49 * sd / (n ^ (1 / 3))
            k = r / h
        
        'Freedman-Diaconis
        ElseIf method = "fd" Then
            iqr = me_quartile_range(data)(1, 2)
            h = 2 * iqr / (n ^ (1 / 3))
            k = r / h
        
        Else
            minBins = 2
            mx = WorksheetFunction.Max(data) + adjust
            mn = WorksheetFunction.Min(data)
        
            r = mx - mn
            
            cMin = Null

            For k = minBins To maxBins
                h = r / k
                
                Dim freq() As Variant
                ReDim freq(0 To k - 1)
                cfOld = 0
                For i = 0 To k - 1
                    lb = mn + i * h
                    ub = lb + h
                    cf = WorksheetFunction.CountIfs(data, "<" & CLng(ub))
                    freq(i) = cf - cfOld
                    cfOld = cf
                Next i
                                
                If method = "shinshim" Then
                    m = n / k
                    s = 0
                    For i = 0 To k - 1
                        s = s + (freq(i) - m) ^ 2
                    Next i
                    v = s / k
                    c = (2 * m - v) / (h ^ 2)
                ElseIf method = "stone" Then
                    s = 0
                    For i = 0 To k - 1
                        s = s + (freq(i) / n) ^ 2
                    Next i
                
                    c = 1 / h * (2 / (n - 1) - (n + 1) / (n - 1) * s)
                ElseIf method = "knuth" Then
                    c1 = n * Log(k) + WorksheetFunction.Gamma(k / 2) - WorksheetFunction.Gamma(n + k / 2)
                    s = 0
                    For i = 0 To k - 1
                        s = s + WorksheetFunction.Gamma(i + 0.5)
                    Next i
                
                    c2 = -k * WorksheetFunction.Gamma(1 / 2) + s
                    c = -1 * (c1 + c2)
                End If
                
                If IsNull(cMin) Then
                    cMin = c
                    kOpt = k
                Else
                    If c < cMin Then
                        cMin = c
                        kOpt = k
                    End If
                End If

            Next k
            
            k = kOpt
                
        End If
    End If
            
    tab_nbins = WorksheetFunction.RoundUp(k, 0)
    

End Function


