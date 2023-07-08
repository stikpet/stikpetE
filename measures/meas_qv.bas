Attribute VB_Name = "meas_qv"
'Created by Peter Stikker
'Companion website: https://PeterStatistics.com
'YouTube channel: https://www.youtube.com/stikpet
'Donations welcome at Patreon: https://www.patreon.com/bePatron?u=19398076

Function me_qv(data As Range, _
            Optional measure = "vr", _
            Optional var1 = 2, _
            Optional var2 = 1, _
            Optional output = "all")
    
    Table = tab_frequency(data)
    freqs = WorksheetFunction.Index(Table, 0, 2)
    k = UBound(Table, 1) - LBound(Table, 1)
    n = WorksheetFunction.Sum(freqs)
    fm = WorksheetFunction.Max(freqs)
    props = WorksheetFunction.Index(Table, 0, 4)
    For i = 2 To k + 1
        props(i, 1) = props(i, 1) / 100
    Next i
    
    If measure = "modvr" Then
        'Modified Variation Ratio
        src = "(Wilcox, 1973, p. 7)"
        lbl = "Wilcox MODVR"
        qv = 0
        For i = 2 To k + 1
            qv = qv + (fm - freqs(i, 1))
        Next i
        qv = qv / (n * (k - 1))
    ElseIf measure = "ranvr" Then
        'Range Variation Ratio
        src = "(Wilcox, 1973, p. 8)"
        lbl = "Wilcox RANVR"
        fl = WorksheetFunction.Min(freqs)
        qv = 1 - (fm - fl) / fm
    ElseIf measure = "avdev" Then
        'Average Deviation
        src = "(Wilcox, 1973, p. 9)"
        lbl = "Wilcox AVDEV"
        qv = 0
        For i = 2 To k + 1
            qv = qv + Abs(freqs(i, 1) - n / k)
        Next i
        qv = 1 - qv / (2 * n / k * (k - 1))
    ElseIf measure = "mndif" Then
        'MNDif
        src = "(Wilcox, 1973, p. 9)"
        lbl = "Wilcox MNDIF"
        mndif = 0
        For i = 2 To k
          For j = i + 1 To k + 1
            mndif = mndif + Abs(freqs(i, 1) - freqs(j, 1))
            Next j
        Next i
        qv = 1 - mndif / (n * (k - 1))
    ElseIf measure = "varnc" Then
        'VarNC
        src = "(Wilcox, 1973, p. 11)"
        lbl = "Wilcox VARNC"
        qv = 0
        For i = 2 To k + 1
            qv = qv + (freqs(i, 1) - n / k) ^ 2
        Next i
        qv = 1 - qv / (n ^ 2 * (k - 1) / k)
    ElseIf measure = "stdev" Then
        src = "(Wilcox, 1973, p. 14)"
        lbl = "Wilcox STDEV"
        For i = 2 To k + 1
            qv = qv + (freqs(i, 1) - n / k) ^ 2
        Next i
        qv = 1 - (qv / ((n - n / k) ^ 2 + (k - 1) * (n / k) ^ 2)) ^ 0.5
    ElseIf measure = "hrel" Then
        'HRel
        src = "(Wilcox, 1973, p. 16)"
        lbl = "Wilcox HREL"
        hrel = 0
        For i = 2 To k + 1
          hrel = hrel + props(i, 1) * WorksheetFunction.Log(props(i, 1), 2)
        Next i
        qv = -hrel / WorksheetFunction.Log(k, 2)
    ElseIf measure = "m1" Then
        src = "(Gibbs & Poston, 1975, p. 471)"
        lbl = "Gibbs-Poston M1"
        qv = 0
        For i = 2 To k + 1
            qv = qv + props(i, 1) ^ 2
        Next i
        qv = 1 - qv
    ElseIf measure = "m2" Then
        'equal to varnc
        src = "(Gibbs & Poston, 1975, p. 472)"
        lbl = "Gibbs-Poston M2"
        qv = 0
        For i = 2 To k + 1
            qv = qv + props(i, 1) ^ 2
        Next i
        qv = (1 - qv) / (1 - 1 / k)
    ElseIf measure = "m3" Then
        src = "(Gibbs & Poston, 1975, p. 472)"
        lbl = "Gibbs-Poston M3"
        pl = WorksheetFunction.Min(props)
        qv = 0
        For i = 2 To k + 1
            qv = qv + props(i, 1) ^ 2
        Next i
        qv = (1 - qv - pl) / (1 - 1 / k - pl)
    ElseIf measure = "m4" Then
        src = "(Gibbs & Poston, 1975, p. 473)"
        lbl = "Gibbs-Poston M4"
        fmean = n / k
        qv = 0
        For i = 2 To k + 1
            qv = qv + Abs(freqs(i, 1) - fmean)
        Next i
        qv = 1 - qv / (2 * n)
    ElseIf measure = "m5" Then
        src = "(Gibbs & Poston, 1975, p. 474)"
        lbl = "Gibbs-Poston M5"
        fmean = n / k
        qv = 0
        For i = 2 To k + 1
            qv = qv + Abs(freqs(i, 1) - fmean)
        Next i
        qv = 1 - qv / (2 * (n - k + 1 - fmean))
    ElseIf measure = "m6" Then
        src = "(Gibbs & Poston, 1975, p. 474)"
        lbl = "Gibbs-Poston M6"
        fmean = n / k
        qv = 0
        For i = 2 To k + 1
            qv = qv + Abs(freqs(i, 1) - fmean)
        Next i
        qv = k * (1 - qv / (2 * n))
    ElseIf measure = "b" Then
        'Kaiser B index
        src = "(Kaiser, 1968, p. 211)"
        lbl = "Kaiser b"
        qv = 1
        For i = 2 To k + 1
            qv = qv * freqs(i, 1) * k / n
        Next i
        qv = 1 - (1 - ((qv) ^ (1 / k)) ^ 2) ^ 0.5
    ElseIf measure = "bd" Then
        'Bulla D
        src = "(Bulla, 1994, p. 169)"
        lbl = "Bulla D"
        o = 0
        For i = 2 To k + 1
            o = o + WorksheetFunction.Min(props(i, 1), 1 / k)
        Next i
        qv = k * (o - 1 / k + (k - 1) / n) / (1 - 1 / k + (k - 1) / n)
    ElseIf measure = "be" Then
        'Bulla e
        src = "(Bulla, 1994, pp. 168-169)"
        lbl = "Bulla E"
        o = 0
        For i = 2 To k + 1
            o = o + WorksheetFunction.Min(props(i, 1), 1 / k)
        Next i
        qv = (o - 1 / k + (k - 1) / n) / (1 - 1 / k + (k - 1) / n)
    ElseIf measure = "bpi" Then
        'Berger-Parker Index
        src = "(Berger & Parker, 1970, p. 1345)"
        lbl = "Berger-Parker D"
        qv = fm / n
    ElseIf measure = "d1" Then
        'Simpson's D
        src = "(Simpson, 1949, p. 688)"
        lbl = "Simpson D"
        qv = 0
        For i = 2 To k + 1
            qv = qv + freqs(i, 1) * (freqs(i, 1) - 1)
        Next i
        qv = qv / (n * (n - 1))
    ElseIf measure = "d2" Then
        'Simpson's D
        src = "(Smith & Wilson, 1996, p. 71)"
        lbl = "Simpson D biased"
        qv = 0
        For i = 2 To k + 1
            qv = qv + (freqs(i, 1) / n) ^ 2
        Next i
        qv = qv
    ElseIf measure = "d3" Then
        'Simpson's D
        src = "(Wikipedia, n.d.)"
        lbl = "Simpson D as diversity"
        qv = 0
        For i = 2 To k + 1
            qv = qv + freqs(i, 1) * (freqs(i, 1) - 1)
        Next i
        qv = 1 - qv / (n * (n - 1))
    ElseIf measure = "d4" Then
        'Simpson's D
        src = "(Berger & Parker, 1970, p. 1345)"
        lbl = "Simpson D as diversity biased"
        qv = 0
        For i = 2 To k + 1
            qv = qv + (freqs(i, 1) / n) ^ 2
        Next i
        qv = 1 - qv
    ElseIf measure = "hd" Then
        'Hill's Diversity
        src = "(Hill, 1973, p. 428)"
        lbl = "Hill Diversity"
        If var1 = 1 Then
            For i = 2 To k + 1
                qv = qv + props(i, 1) * Log(props(i, 1))
            Next i
            qv = Exp(-1 * qv)
        Else
            For i = 2 To k + 1
                qv = qv + props(i, 1) ^ var1
            Next i
            qv = 1 / (qv ^ (var1 - 1))
        End If
    ElseIf measure = "he" Then
        'Hill's Evenness
        src = "(Hill, 1973, p. 429)"
        lbl = "Hill Evenness"
        qv = me_qv(data, "hd", var1, , "value") / me_qv(data, "hd", var2, , "value")
    ElseIf measure = "hi" Then
        'Heip Index
        src = "(Heip, 1974, p. 555)"
        lbl = "Heip Evenness"
        h = 0
        For i = 2 To k + 1
            h = h + props(i, 1) * Log(props(i, 1))
        Next i
        h = -1 * h
        qv = (Exp(h) - 1) / (k - 1)
    ElseIf measure = "j" Then
        'Pielou J
        src = "(Pielou, 1966, p. 141)"
        lbl = "Pielou J"
        h = 0
        For i = 2 To k + 1
            h = h + props(i, 1) * Log(props(i, 1))
        Next i
        h = -1 * h
        qv = h / Log(k)
    ElseIf measure = "si" Then
        'Sheldon Index
        src = "(Sheldon, 1969, p. 467)"
        lbl = "Sheldon Evenness"
        h = 0
        For i = 2 To k + 1
            h = h + props(i, 1) * Log(props(i, 1))
        Next i
        h = -1 * h
        qv = Exp(h) / k
    ElseIf measure = "sw1" Then
        'Smith and Wilson Index 1
        src = "(Smith & Wilson, 1996, p. 71)"
        lbl = "Smith-Wilson Evenness Index 1"
        d = 0
        For i = 2 To k + 1
            d = d + props(i, 1) ^ 2
        Next i
        qv = (1 - d) / (1 - 1 / k)
    ElseIf measure = "sw2" Then
        'Smith and Wilson Index 2
        src = "(Smith & Wilson, 1996, p. 71)"
        lbl = "Smith-Wilson Evenness Index 2"
        d = 0
        For i = 2 To k + 1
            d = d + props(i, 1) ^ 2
        Next i
        qv = -Log(d) / Log(k)
    ElseIf measure = "sw3" Then
        'Smith and Wilson Index 3
        src = "(Smith & Wilson, 1996, p. 71)"
        lbl = "Smith-Wilson Evenness Index 3"
        d = 0
        For i = 2 To k + 1
            d = d + props(i, 1) ^ 2
        Next i
        qv = 1 / (d * k)
    ElseIf measure = "swe" Then
        'Shannon-Weaver Entropy
        src = "(Shannon & Weaver, 1949, p. 20)"
        lbl = "Shannon-Weaver Entropy"
        h = 0
        For i = 2 To k + 1
            h = h + props(i, 1) * Log(props(i, 1))
        Next i
        qv = -1 * h
    ElseIf measure = "re" Then
        'Renyi Entropy
        src = "(Renyi, 1961, p. 549)"
        lbl = "Reneyi Entropy"
        qv = 0
        For i = 2 To k + 1
            qv = qv + props(i, 1) ^ var1
        Next i
        qv = 1 / (1 - var1) * WorksheetFunction.Log(qv, 2)
    ElseIf measure = "vr" Then
        'Variation Ratio
        src = "(Freeman, 1965)"
        lbl = "Freeman Variation Ratio"
        pm = fm / n
        qv = 1 - pm
    End If
    
    'result to show
    If output = "value" Then
        me_qv = qv
    ElseIf output = "measure" Then
        me_qv = lbl
    ElseIf output = "source" Then
        me_qv = src
    ElseIf output = "all" Then
        Dim res(1 To 2, 1 To 3)
        res(1, 1) = "value"
        res(1, 2) = "measure"
        res(1, 3) = "source"
        res(2, 1) = qv
        res(2, 2) = lbl
        res(2, 3) = src
        me_qv = res
    End If

End Function
