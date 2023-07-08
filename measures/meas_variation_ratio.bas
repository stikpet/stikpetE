Attribute VB_Name = "meas_variation_ratio"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Function me_variation_ratio(data As Range)
Attribute me_variation_ratio.VB_Description = "Variation Ratio"
Attribute me_variation_ratio.VB_ProcData.VB_Invoke_Func = " \n14"
'uses the me_mode function

    modeInfo = me_mode(data)
    
    maxFreq = modeInfo(2, 2)
    modes = modeInfo(2, 1)
    
    If modes = "none" Then
        nModes = 0
        me_variation_ratio = "no mode in data, so also no variation ratio"
    Else
        nModes = Len(modes) - Len(WorksheetFunction.Substitute(modes, ",", "")) + 1
        
        n = WorksheetFunction.CountA(data)
    
        vr = 1 - nModes * maxFreq / n
        
        me_variation_ratio = vr
    
    End If

End Function
