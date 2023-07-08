Attribute VB_Name = "dis_multinomial"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

'Used for ts_trinomial_os

Function di_mpmf(counts, probs)
    If TypeName(counts) = "Range" Then
        nr = counts.Rows.Count
    Else
        nr = UBound(counts, 1) - LBound(counts, 1) + 1
    End If
    
    Dim freq() As Double
    Dim prs() As Double
    ReDim freq(1 To nr, 1 To 1)
    ReDim prs(1 To nr, 1 To 1)
    
    
    For i = 1 To nr
        freq(i, 1) = counts(i, 1)
        prs(i, 1) = probs(i, 1)
    Next i

    k = UBound(freq, 1) - LBound(freq, 1) + 1
    
    den = 1
    n = 0
    For i = 1 To k
        den = den * WorksheetFunction.Gamma(freq(i, 1) + 1)
        n = n + freq(i, 1)
    Next i
    
    factor1 = WorksheetFunction.Gamma(1 + n) / den
    
    factor2 = 1
    For i = 1 To k
        factor2 = factor2 * probs(i, 1) ^ (freq(i, 1))
    Next i
    
    pVal = factor1 * factor2
    
    di_mpmf = pVal
    
    
End Function

