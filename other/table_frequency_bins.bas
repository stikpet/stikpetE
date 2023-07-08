Attribute VB_Name = "table_frequency_bins"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

'this function makes use of the tab_nbins() function

Function tab_frequency_bins(data As Range, _
                    Optional nbins = Null, _
                    Optional bins As Range, _
                    Optional incl_lower = True, _
                    Optional adjust = 1)
    
    If bins Is Nothing Then
        If IsNull(nbins) Then
            k = tab_nbins(data)
        Else
            k = tab_nbins(data, nbins)
        End If
    Else
        k = bins.Rows.Count
    End If
    
    'determine minimum and maximum
    mx = WorksheetFunction.Max(data)
    mn = WorksheetFunction.Min(data)
    
    'increase maximimum if to include the lower bound
    If incl_lower Then
        mx = mx + adjust
    Else
    'decrease minimum if to include the upper bound
        mn = mn - adjust
    End If
    
    'determine range and width
    r = mx - mn
    h = r / k
    
    'create the table
    Dim freq As Variant
    ReDim freq(0 To k, 1 To 4)
    'set column titles
    freq(0, 1) = "lower bound"
    freq(0, 2) = "upper bound"
    freq(0, 3) = "frequency"
    freq(0, 4) = "frequency density"
    'fill the rows
    For i = 1 To k
        If bins Is Nothing Then
            lb = mn + (i - 1) * h
            ub = lb + h
        Else
            lb = bins(i, 1)
            ub = bins(i, 2)
        End If
        
        If incl_lower Then
            f = WorksheetFunction.CountIf(data, "<" & CLng(ub)) - WorksheetFunction.CountIf(data, "<" & CLng(lb))
        Else
            f = WorksheetFunction.CountIf(data, "<=" & CLng(ub)) - WorksheetFunction.CountIf(data, "<=" & CLng(lb))
        End If
        fd = f / (ub - lb)
        freq(i, 1) = lb
        freq(i, 2) = ub
        freq(i, 3) = f
        freq(i, 4) = fd
    Next i
    
    tab_frequency_bins = freq

End Function
