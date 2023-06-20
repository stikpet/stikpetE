Attribute VB_Name = "meas_quartile_range"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Public Sub me_quartile_range_addHelp()
Application.MacroOptions _
    Macro:="me_quartile_range", _
    Description:="Quartile Range", _
    category:=14, _
    ArgumentDescriptions:=Array( _
        "vertical specific range with data", _
        "optional vertical range with labels in order if data is non-numeric.", _
        "optional which measure to calculate. Either iqr, siqr, qd, or mqr", _
        "optional which method to use to calculate quartiles indexMethod can be set to Can be set to inclusive, exclusive, sas1, sas2, sas3, sas5, sas4, ms, lohninger, hl2, hl1, excel, pd2, pd3, pd4, hf3b, pd5, hf8, hf9, maple2", _
        "optional which output to show. Either all (default) or value")
               
End Sub

Function me_quartile_range(data As range, _
                            Optional levels As range, _
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


