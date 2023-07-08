Attribute VB_Name = "meas_consensus"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Function me_consensus(data As Range, Optional levels As Range)
Attribute me_consensus.VB_Description = "Concensus"
Attribute me_consensus.VB_ProcData.VB_Invoke_Func = " \n14"
'Uses the tab_frequency function

    'get data as numeric values
    If levels Is Nothing Then
        dataN = he_range_to_num_array(data)
    Else
        dataN = he_replace(data, levels)
    End If
    
    'sort the numeric values
    dataN = he_sort(dataN)
    
    'determining the rows
    nrRows = UBound(dataN, 1) - LBound(dataN, 1) + 1
    
    freq = tab_frequency(data, levels)
    
    k = UBound(freq, 1) - LBound(freq, 1)
    n = 0
    m = 0
    For i = 1 To k
        n = n + freq(i, 2)
        m = m + i * freq(i, 2)
    Next i
    m = m / n
    
    cns = 1
    For i = 1 To k
        cns = cns + freq(i, 2) / n * WorksheetFunction.Log(1 - Abs(i - m) / (k - 1), 2)
    Next i
    
    me_consensus = cns
    

End Function

