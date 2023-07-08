Attribute VB_Name = "meas_quartile_range"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076


Function me_quartile_range(data As Range, _
                            Optional levels As Range, _
                            Optional measure = "iqr", _
                            Optional method = "cdf", _
                            Optional output = "all")


    If levels Is Nothing Then
        qs = me_quartiles(data, , method)
    Else
        qs = me_quartiles(data, levels, method)
    End If
    
    q1 = qs(1, 0)
    q3 = qs(1, 1)
    
    If measure = "iqr" Then
        r = q3 - q1
        If method = "tukey" Or method = "inclusive" Or method = "tukey" Or method = "vining" Or method = "hinges" Then
            rName = "Hspread"
        Else
            rName = "IQR"
        End If
    ElseIf measure = "siqr" Or measure = "qd" Then
        r = (q3 - q1) / 2
        rName = "SIQR"
    ElseIf measure = "mqr" Then
        r = (q3 + q1) / 2
        rName = "MQR"
    End If
    
    If output = "value" Then
        me_quartile_range = r
    Else
        Dim res(0 To 1, 0 To 2)
        res(0, 0) = "Q1"
        res(0, 1) = "Q3"
        res(0, 2) = rName
        res(1, 0) = q1
        res(1, 1) = q3
        res(1, 2) = r
        me_quartile_range = res
    End If
        

End Function


