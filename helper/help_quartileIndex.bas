Attribute VB_Name = "help_quartileIndex"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076
Function he_quartileIndex(data, indexMethod, Optional q1Frac = "linear", Optional q1Int = "int", Optional q3Frac = "linear", Optional q3Int = "int")

    n = UBound(data, 1) - LBound(data, 1) + 1
    iqs = he_quartileIndexing(data, indexMethod)
    iq1 = iqs(0)
    iq3 = iqs(1)

    If Round(iq1, 0) = iq1 Then
        ' index is integer
        If q1Int = "int" Then
            q1 = iq1
        ElseIf q1Int = "midpoint" Then
            q1 = iq1 + 1 / 2
        End If
    Else
        ' index has fraction
        If q1Frac = "linear" Then
            q1 = iq1
        ElseIf q1Frac = "down" Then
            q1 = WorksheetFunction.RoundDown(iq1, 0)
        ElseIf q1Frac = "up" Then
            q1 = WorksheetFunction.RoundUp(iq1, 0)
        ElseIf q1Frac = "bankers" Then
            q1 = Round(iq1, 0)
        ElseIf q1Frac = "nearest" Then
            q1 = WorksheetFunction.RoundDown(iq1 + 0.5, 0)
        ElseIf q1Frac = "halfdown" Then
            If iq1 + 0.5 = Round(iq1 + 0.5, 0) Then
                q1 = WorksheetFunction.RoundDown(iq1, 0)
            Else
                q1 = Round(iq1, 0)
            End If
        ElseIf q1Frac = "midpoint" Then
            q1 = (WorksheetFunction.RoundDown(iq1, 0) + WorksheetFunction.RoundUp(iq1, 0)) / 2
        End If
    End If
    
    q1i = q1
    q1iLow = WorksheetFunction.RoundDown(q1i, 0)
    q1iHigh = WorksheetFunction.RoundUp(q1i, 0)

    If q1iLow = q1iHigh Then
        q1 = data(WorksheetFunction.RoundDown(q1iLow - 1, 0))
    Else
        'Linear interpolation:
        q1 = data(WorksheetFunction.RoundDown(q1iLow - 1, 0)) + (q1i - q1iLow) / (q1iHigh - q1iLow) * (data(WorksheetFunction.RoundDown(q1iHigh - 1, 0)) - data(WorksheetFunction.RoundDown(q1iLow - 1, 0)))
    End If
    
    If Round(iq3, 0) = iq3 Then
        ' index is integer
        If q3Int = "int" Then
            q3 = iq3
        ElseIf q3Int = "midpoint" Then
            q3 = iq3 + 1 / 2
        End If
    Else
        ' index has fraction
        If q3Frac = "linear" Then
            q3 = iq3
        ElseIf q3Frac = "down" Then
            q3 = WorksheetFunction.RoundDown(iq3, 0)
        ElseIf q3Frac = "up" Then
            q3 = WorksheetFunction.RoundUp(iq3, 0)
        ElseIf q3Frac = "bankers" Then
            q3 = Round(iq3, 0)
        ElseIf q3Frac = "nearest" Then
            q3 = WorksheetFunction.RoundDown(iq3 + 0.5, 0)
        ElseIf q3Frac = "halfdown" Then
            If iq3 + 0.5 = Round(iq3 + 0.5, 0) Then
                q3 = WorksheetFunction.RoundDown(iq3, 0)
            Else
                q3 = Round(iq3, 0)
            End If
        ElseIf q3Frac = "midpoint" Then
            q3 = (WorksheetFunction.RoundDown(iq3, 0) + WorksheetFunction.RoundUp(iq3, 0)) / 2
        End If
    End If
                
    q3i = q3
    q3iLow = WorksheetFunction.RoundDown(q3i, 0)
    q3iHigh = WorksheetFunction.RoundUp(q3i, 0)

    If q3iLow = q3iHigh Then
        q3 = data(WorksheetFunction.RoundDown(q3iLow - 1, 0))
    Else
        'Linear interpolation:
        q3 = data(WorksheetFunction.RoundDown(q3iLow - 1, 0)) + (q3i - q3iLow) / (q3iHigh - q3iLow) * (data(WorksheetFunction.RoundDown(q3iHigh - 1, 0)) - data(WorksheetFunction.RoundDown(q3iLow - 1, 0)))
    End If
    
    Dim results(0 To 1) As Double
    results(0) = q1
    results(1) = q3
    
    he_quartileIndex = results

End Function

