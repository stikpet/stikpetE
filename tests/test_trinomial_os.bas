Attribute VB_Name = "test_trinomial_os"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

'uses the helper functions he_range_to_num_array() and he_replace()
'uses the function di_mpmf()

Function ts_trinomial_os(data As Range, _
                        Optional levels As Range, _
                        Optional mu = Null, _
                        Optional output = "all")
    
    'get data as numeric values
    If levels Is Nothing Then
        dataN = he_range_to_num_array(data)
    Else
        dataN = he_replace(data, levels)
    End If
        
    'set mu to midrange if not provided
    If IsNull(mu) Then
        mu = (WorksheetFunction.Min(dataN) + WorksheetFunction.Max(dataN)) / 2
    End If
    
    If output = "mu" Then
        ts_trinomial_os = mu
    Else
    
        nPos = 0
        nNeg = 0
        nNul = 0
        For Each i In dataN
            If i > mu Then
                nPos = nPos + 1
            ElseIf i < mu Then
                nNeg = nNeg + 1
            ElseIf i = mu Then
                nNul = nNul + 1
            End If
        Next i
        nd = Abs(nPos - nNeg)
        n = nPos + nNeg + nNul
        
        p0 = nNul / n
        p1 = (1 - p0) / 2
        
        Dim probs(1 To 3, 1 To 1)
        probs(1, 1) = p1
        probs(2, 1) = p1
        probs(3, 1) = p0
        
        sig = 0
        Dim ns(1 To 3, 1 To 1)
        For Z = nd To n
           For i = 0 To WorksheetFunction.RoundDown((n - Z) / 2, 0)
                ns(1, 1) = i
                ns(2, 1) = i + Z
                ns(3, 1) = n - i - (i + Z)
                sig = sig + di_mpmf(ns, probs)
            Next i
        Next Z
        
        pValue = sig * 2
        If pValue > 1 Then
            pValue = 1
        End If
        
        If output = "all" Then
            Dim results(1 To 2, 1 To 6)
            results(1, 1) = "mu"
            results(1, 2) = "n-pos."
            results(1, 3) = "n-neg."
            results(1, 4) = "n-tied."
            results(1, 5) = "p-value"
            results(1, 6) = "test"
            
            results(2, 1) = mu
            results(2, 2) = nPos
            results(2, 3) = nNeg
            results(2, 4) = nNul
            results(2, 5) = pValue
            results(2, 6) = "one-sample trinomial"
            
            ts_trinomial_os = results
            
        ElseIf output = "pvalue" Then
            ts_trinomial_os = pValue
        End If
            
    End If
    
End Function

