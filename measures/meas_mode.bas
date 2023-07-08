Attribute VB_Name = "meas_mode"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076


Function me_mode(data As Range, Optional allEq = "none", Optional output = "all")
Attribute me_mode.VB_Description = "Mode"
Attribute me_mode.VB_ProcData.VB_Invoke_Func = " \n14"
'makes use of the tab_frequency function
    
    freq = tab_frequency(data)
    
    fMode = 0
    k = UBound(freq, 1) - LBound(freq, 1)
    
    For i = 1 To k
        If freq(i, 2) > fMode Then
            fMode = freq(i, 2)
            modes = freq(i, 1)
            nModes = 1
        ElseIf freq(i, 2) = fMode Then
            modes = modes & ", " & freq(i, 1)
            nModes = nModes + 1
        End If
    Next i

    If nModes = k And allEq = "none" Then
        modes = "none"
        fMode = "none"
    End If
    
    'results
    If output = "all" Then
        Dim res(1 To 2, 1 To 2)
        res(1, 1) = "mode"
        res(1, 2) = "mode frequency"
        res(2, 1) = modes
        res(2, 2) = fMode
        me_mode = res
    ElseIf output = "mode" Then
        me_mode = modes
    ElseIf output = "freq" Then
        me_mode = fMode
    End If
        
End Function
