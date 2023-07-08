Attribute VB_Name = "meas_quartiles"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Function me_quartiles(data As Range, Optional levels As Range, Optional method = "own", Optional indexMethod = "sas1", Optional q1Frac = "linear", Optional q1Int = "int", Optional q3Frac = "linear", Optional q3Int = "int")

    'get data as numeric values
    If levels Is Nothing Then
        dataN = he_range_to_num_array(data)
    Else
        dataN = he_replace(data, levels)
    End If
    
    'sort the numeric values
    dataN = he_sort(dataN)
    
    'Set to default method name
    Select Case method
        Case "inclusive", "tukey", "vining", "hinges"
            method = "inclusive"
        Case "exclusive", "jf"
            method = "exclusive"
        Case "cdf", "sas5", "hf2", "averaged_inverted_cdf", "r2"
            method = "sas5"
        Case "sas4", "minitab", "hf6", "weibull", "maple5", "r6"
            method = "sas4"
        Case "excel", "hf7", "pd1", "linear", "gumbel", "maple6", "r7"
            method = "excel"
        Case "sas1", "parzen", "hf4", "interpolated_inverted_cdf", "maple3", "r4"
            method = "sas1"
        Case "sas2", "hf3", "r3"
            method = "sas2"
        Case "sas3", "hf1", "inverted_cdf", "maple1", "r1"
            method = "sas3"
        Case "hf3b", "closest_observation"
            method = "hf3b"
        Case "hl2", "hazen", "hf5", "maple4"
            method = "hl2"
        Case "np", "midpoint", "pd5"
            method = "pd5"
        Case "hf8", "median_unbiased", "maple7", "r8"
            method = "hf8"
        Case "hf9", "normal_unbiased", "maple8", "r9"
            method = "hf9"
        Case "pd2", "lower"
            method = "pd2"
        Case "pd3", "higher"
            method = "pd3"
        Case "pd4", "nearest"
            method = "pd4"
    End Select
    
    'settings
    Dim settings(0 To 4) As String
    settings(0) = indexMethod
    settings(1) = q1Frac
    settings(2) = q1Int
    settings(3) = q3Frac
    settings(4) = q3Int
    
    If method = "inclusive" Then
        settings(0) = "inclusive"
        settings(1) = "linear"
        settings(2) = "int"
        settings(3) = "linear"
        settings(4) = "int"
    ElseIf method = "exclusive" Then
        settings(0) = "exclusive"
        settings(1) = "linear"
        settings(2) = "int"
        settings(3) = "linear"
        settings(4) = "int"
    ElseIf method = "sas1" Then
        settings(0) = "sas1"
        settings(1) = "linear"
        settings(2) = "int"
        settings(3) = "linear"
        settings(4) = "int"
    ElseIf method = "sas2" Then
        settings(0) = "sas1"
        settings(1) = "bankers"
        settings(2) = "int"
        settings(3) = "bankers"
        settings(4) = "int"
    ElseIf method = "sas3" Then
        settings(0) = "sas1"
        settings(1) = "up"
        settings(2) = "int"
        settings(3) = "up"
        settings(4) = "int"
    ElseIf method = "sas5" Then
        settings(0) = "sas1"
        settings(1) = "up"
        settings(2) = "midpoint"
        settings(3) = "up"
        settings(4) = "midpoint"
    ElseIf method = "sas4" Then
        settings(0) = "sas4"
        settings(1) = "linear"
        settings(2) = "int"
        settings(3) = "linear"
        settings(4) = "int"
    ElseIf method = "ms" Then
        settings(0) = "sas4"
        settings(1) = "nearest"
        settings(2) = "int"
        settings(3) = "halfdown"
        settings(4) = "int"
    ElseIf method = "lohninger" Then
        settings(0) = "sas4"
        settings(1) = "nearest"
        settings(2) = "int"
        settings(3) = "nearest"
        settings(4) = "int"
    ElseIf method = "hl2" Then
        settings(0) = "hl"
        settings(1) = "linear"
        settings(2) = "int"
        settings(3) = "linear"
        settings(4) = "int"
    ElseIf method = "hl1" Then
        settings(0) = "hl"
        settings(1) = "midpoint"
        settings(2) = "int"
        settings(3) = "midpoint"
        settings(4) = "int"
    ElseIf method = "excel" Then
        settings(0) = "excel"
        settings(1) = "linear"
        settings(2) = "int"
        settings(3) = "linear"
        settings(4) = "int"
    ElseIf method = "pd2" Then
        settings(0) = "excel"
        settings(1) = "down"
        settings(2) = "int"
        settings(3) = "down"
        settings(4) = "int"
    ElseIf method = "pd3" Then
        settings(0) = "excel"
        settings(1) = "up"
        settings(2) = "int"
        settings(3) = "up"
        settings(4) = "int"
    ElseIf method = "pd4" Then
        settings(0) = "excel"
        settings(1) = "halfdown"
        settings(2) = "int"
        settings(3) = "nearest"
        settings(4) = "int"
    ElseIf method = "hf3b" Then
        settings(0) = "sas1"
        settings(1) = "nearest"
        settings(2) = "int"
        settings(3) = "halfdown"
        settings(4) = "int"
    ElseIf method = "pd5" Then
        settings(0) = "excel"
        settings(1) = "midpoint"
        settings(2) = "int"
        settings(3) = "midpoint"
        settings(4) = "int"
    ElseIf method = "hf8" Then
        settings(0) = "hf8"
        settings(1) = "linear"
        settings(2) = "int"
        settings(3) = "linear"
        settings(4) = "int"
    ElseIf method = "hf9" Then
        settings(0) = "hf9"
        settings(1) = "linear"
        settings(2) = "int"
        settings(3) = "linear"
        settings(4) = "int"
    ElseIf method = "maple2" Then
        settings(0) = "hl"
        settings(1) = "down"
        settings(2) = "int"
        settings(3) = "down"
        settings(4) = "int"
    End If
    
    iqs = he_quartileIndex(dataN, settings(0), settings(1), settings(2), settings(3), settings(4))
    q1 = iqs(0)
    q3 = iqs(1)
    
    'find the text representatives
    If levels Is Nothing Then
    
        Dim results1(0 To 1, 0 To 1) As Variant
        results1(0, 0) = "q1"
        results1(0, 1) = "q3"
        results1(1, 0) = q1
        results1(1, 1) = q3
        me_quartiles = results1
    
    Else
        If q1 = Round(q1, 0) Then
            q1T = levels(q1, 1)

        Else
            q1T = "between " + levels(WorksheetFunction.RoundDown(q1, 0), 1) + " and " + levels(WorksheetFunction.RoundUp(q1, 0), 1)
        End If

        If q3 = Round(q3, 0) Then
            q3T = levels(q3, 1)

        Else
            q3T = "between " + levels(WorksheetFunction.RoundDown(q3, 0), 1) + " and " + levels(WorksheetFunction.RoundUp(q3, 0), 1)
        End If
        
        
        Dim results2(0 To 1, 0 To 3) As Variant
        results2(0, 0) = "q1"
        results2(0, 1) = "q3"
        results2(0, 2) = "q1-Text"
        results2(0, 3) = "q3-Text"
        results2(1, 0) = q1
        results2(1, 1) = q3
        results2(1, 2) = q1T
        results2(1, 3) = q3T
        
        me_quartiles = results2
        
    End If
        

End Function


