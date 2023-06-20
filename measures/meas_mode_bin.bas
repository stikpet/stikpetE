Attribute VB_Name = "meas_mode_bin"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Public Sub me_mode_bin_addHelp()
Application.MacroOptions _
    Macro:="me_mode_bin", _
    Description:="Mode for Binned Data", _
    category:=14, _
    ArgumentDescriptions:=Array( _
        "range with three columns containing the lower bounds, upper bounds and frequencies", _
        "optional to indicate what to do if all frequencies are equal, either none (default) or all", _
        "optional to indicate method to calculate specific mode value, either none (default), midpoint or quadratic", _
        "optional to indicate the output to show, either all (default), mode or fd")
               
End Sub

Function me_mode_bin(binData As range, Optional allEq = "none", Optional value = "none", Optional output = "all")
'binData should be a range with three columns: lower bound, upper bound and frequency
    
    Dim fd() As Double
    Dim iMode() As Integer
    
    
    k = binData.Rows.Count
    
    'determine the frequency densities, the modal frequency density
    modeFD = 0
    nModes = 0
    ReDim fd(1 To k) As Double
    ReDim iMode(1 To k) As Integer
    For i = 1 To k
        fd(i) = binData.Cells(i, 3) / (binData.Cells(i, 2) - binData.Cells(i, 1))
        
        If fd(i) > modeFD Then
            Mode = binData.Cells(i, 1) & " < " & binData.Cells(i, 2)
            nModes = 1
            iMode(nModes) = i
            modeFD = fd(i)
        ElseIf fd(i) = modeFD Then
            Mode = Mode & ", " & binData.Cells(i, 1) & " < " & binData.Cells(i, 2)
            nModes = nModes + 1
            iMode(nModes) = i
        End If
        
    Next i
    
    If nModes = k And allEq = "none" Then
        Mode = "none"
        modeFD = "none"
    Else
        If value = "midpoint" Then
            Mode = (binData.Cells(iMode(1), 2) + binData.Cells(iMode(1), 1)) / 2
            If nModes > 1 Then
                For i = 2 To nModes
                    Mode = Mode & ", " & (binData.Cells(iMode(i), 2) + binData.Cells(iMode(i), 1)) / 2
                Next i
            End If
        
        ElseIf value = "quadratic" Then
            d1 = modeFD
            d2 = modeFD - fd(2)
            Mode = binData.Cells(iMode(1), 1) + d1 / (d1 + d2) * (binData.Cells(iMode(1), 2) - binData.Cells(iMode(1), 1))
            If nModes > 1 Then
                For i = 2 To nModes
                    d1 = modeFD - fd(iMode(i) - 1)
                    
                    If i = k Then
                        d2 = modeFD
                    Else
                        d2 = modeFD - fd(iMode(i) + 1)
                    End If
                    
                    Mode = Mode & ", " & binData.Cells(iMode(i), 1) + d1 / (d1 + d2) * (binData.Cells(iMode(i), 2) - binData.Cells(iMode(i), 1))
                Next i
            End If
        End If
    End If
    
    If output = "mode" Then
        me_mode_bin = Mode
    ElseIf output = "fd" Then
        me_mode_bin = modeFD
    Else
        Dim res(1 To 2, 1 To 2)
        res(1, 1) = "mode"
        res(1, 2) = "modeFD"
        res(2, 1) = Mode
        res(2, 2) = modeFD
        me_mode_bin = res
    End If

End Function
