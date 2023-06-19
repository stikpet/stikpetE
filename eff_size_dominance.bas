Attribute VB_Name = "eff_size_dominance"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Public Sub es_dominance_addHelp()
Application.MacroOptions _
    Macro:="es_dominance", _
    Description:="Dominance and a Vargha-Delaney A like effect size measure", _
    category:=14, _
    ArgumentDescriptions:=Array( _
        "vertical specific range with data", _
        "optional vertical range with labels in order if data is non-numeric.", _
        "optional parameter to set the hypothesized median. If not used the midrange is used", _
        "optional to set what output to show. Either dominance (default), vda, domValue, vdaValue, or mu")
                
End Sub


Function es_dominance(data As Range, _
                    Optional levels As Range, _
                    Optional mu = "none", _
                    Optional output = "dominance")
                    
'get data as numeric values
    If levels Is Nothing Then
        dataN = he_range_to_num_array(data)
    Else
        dataN = he_replace(data, levels)
    End If
    
    'sort the numeric values
    dataN = he_sort(dataN)
    
    'set mu to midrange if not provided
    If mu = "none" Then
        mu = (WorksheetFunction.Min(dataN) + WorksheetFunction.Max(dataN)) / 2
    End If
    
    'sample size
    n = UBound(dataN, 1) - LBound(dataN, 1) + 1
    
    'proportion of scores above mu and below
    pPlus = 0
    pMin = 0
    For i = 0 To n - 1
        If dataN(i) < mu Then
            pMin = pMin + 1
        ElseIf dataN(i) > mu Then
            pPlus = pPlus + 1
        End If
    Next i
    
    pPlus = pPlus / n
    pMin = pMin / n
    
    'dominance is simply the difference
    dominance = pPlus - pMin
    
    'the VDA like effect size
    vda = (dominance + 1) / 2
    
    'results
    Dim res(1 To 2, 1 To 2)
    If output = "vdaValue" Then
        es_dominance = vda
    ElseIf output = "domValue" Then
        es_dominance = dominance
    ElseIf output = "mu" Then
        es_dominance = mu
    ElseIf output = "dominance" Then
        
        res(1, 1) = "mu"
        res(1, 2) = "dominance"
        res(2, 1) = mu
        res(2, 2) = dominance
        
        es_dominance = res
        
    ElseIf output = "vda" Then
        
        res(1, 1) = "mu"
        res(1, 2) = "VDA-like"
        res(2, 1) = mu
        res(2, 2) = vda
        
        es_dominance = res
    End If

End Function
