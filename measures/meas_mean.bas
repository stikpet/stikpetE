Attribute VB_Name = "meas_mean"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Function me_mean(data As Range, Optional version = "arithmetic", Optional trimProp = 0.1, Optional trimFrac = "down")
Attribute me_mean.VB_Description = "Mean"
Attribute me_mean.VB_ProcData.VB_Invoke_Func = " \n14"

    Select Case version
    Case "arithmetic"
        res = WorksheetFunction.Average(data)
        
    Case "winsorized", "trimmed", "windsor", "truncated"
        
        dataA = he_range_to_num_array(data)
        dataS = he_sort(dataA)
        n = WorksheetFunction.Count(data)
        ntlow = n * trimProp / 2
        nlow = WorksheetFunction.RoundDown(ntlow, 0)
        
        If version = "winsorized" Then
            s = 0
            For i = nlow To n - nlow - 1
                s = s + dataS(i)
            Next i
            s = s + nlow * (dataS(nlow) + dataS(n - nlow - 1))
            res = s / n
            
        Else
            If trimFrac = "down" Then
                s = 0
                For i = nlow To n - nlow - 1
                    s = s + dataS(i)
                Next i
                res = s / (n - 2 * nlow)
            ElseIf trimFrac = "prop" Then
                fr = ntlow - nlow
                s = 0
                For i = nlow + 1 To n - nlow - 2
                    s = s + dataS(i)
                Next i
                res = ((dataS(nlow) + dataS(n - nlow - 1)) * (1 - fr) + s) / (n - 2 * ntlow)
            ElseIf trimFrac = "linear" Then
                p1 = nlow * 2 / n
                p2 = (nlow + 1) * 2 / n
                s = 0
                For i = nlow To n - nlow - 1
                    s = s + dataS(i)
                Next i
                m1 = s / (n - 2 * nlow)
                m2 = (s - dataS(nlow) - dataS(n - nlow - 1)) / (n - 2 * nlow - 2)
                res = (trimProp - p1) / (p2 - p1) * (m2 - m1) + m1
            End If
        End If
    
    Case "olympic"
        n = WorksheetFunction.Count(data)
        Max = WorksheetFunction.Max(data)
        Min = WorksheetFunction.Min(data)
        s = WorksheetFunction.Sum(data)
        res = (s - Max - Min) / (n - 2)
    
    Case "geometric"
        s = 0
        For Each Cell In data
            If Not IsEmpty(Cell.value) And WorksheetFunction.IsNumber(Cell.value) Then
                s = s + Log(Cell.value)
            End If
        Next Cell
        res = Exp(s)
    
    Case "harmonic"
        n = WorksheetFunction.Count(data)
        For Each Cell In data
            If Not IsEmpty(Cell.value) And WorksheetFunction.IsNumber(Cell.value) Then
                s = s + 1 / Cell.value
            End If
        Next Cell
        res = n / s
    
    Case "midrange"
        Max = WorksheetFunction.Max(data)
        Min = WorksheetFunction.Min(data)
        res = (Max + Min) / 2
                              
    End Select
    
    me_mean = res

End Function
