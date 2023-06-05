Attribute VB_Name = "convert_es"
Function es_convert(es, from, target, Optional ex1 = Null, Optional ex2 = Null)

'COHEN d
  'Cohen d one-sample target Cohen d
  If (from = "cohendos" And target = "cohend") Then
    res = es * Sqr(2)
  
  'Cohen d target Odds Ratio
  ElseIf (from = "cohend" And target = "or") Then
    'Chinn (2000, p. 3129)
    If (ex1 = "chinn") Then
      res = Exp(1.81 * es)
    Else
      'Borenstein et. al (2009, p. 3)
      res = Exp(es * Pi / Sqr(3))
    End If

  'COHEN F
  'Cohen f target eta squared
  ElseIf (from = "cohenf" And target = "etasq") Then
    res = es ^ 2 / (1 + es ^ 2)
  
  'COHEN h'
  'Cohen h' target Cohen h
  ElseIf (from = "cohenhos" And target = "cohenh") Then
    res = es * Sqr(2)

  'COHEN w
  'Cohen w target Contingency Coefficient
  ElseIf (from = "cohenw" And target = "cc") Then
    res = Sqr(es ^ 2 / (1 + es ^ 2))


  'CONTINGENCY COEFFICIENT
  'Contingency Coefficient target Cohen w
  ElseIf (from = "cc" And target = "cohenw") Then
    res = Sqr(es ^ 2 / (1 - es ^ 2))


  'CRAMÉR V
  'Cramer's v GoF target Cohen w
  ElseIf (from = "cramervgof" And target = "cohenw") Then
    res = es * Sqr(ex1 - 1)


  'EPSILON SQUARED
  'Epsilon squared target Eta squared
  ElseIf (from = "epsilonsq" And target = "etasq") Then
    res = 1 - (1 - es) * (ex1 - ex2) / (ex1 - 1)
  'Epsilon squared target Omega squared
  ElseIf (from = "epsilonsq" And target = "omegasq") Then
    res = es * (1 - ex1 / (ex2 + ex1))


  'ETA SQUARED
  'Eta squared target Cohen f
  ElseIf (from = "etasq" And target = "cohenf") Then
    res = Sqr(es / (1 - es))

  'Eta squared target Epsilon Squared
  ElseIf (from = "etasq" And target = "epsilonsq") Then
    res = (ex1 * es - ex2 + (1 - es)) / (ex1 - ex2)

  'JOHNSTON-BERRY-MIELKE E
  'Johnston-Berry-Mielke E to Cohen w
  ElseIf (from = "jbme" And target = "cohenw") Then
    res = Sqr(es * (1 - ex1) / (ex1))

  'ODDS RATIO
  'Odds Ratio target Cohen d (Chinn, 2000, p. 3129)
  ElseIf (from = "or" And target = "cohend") Then
    If (ex1 = "chinn") Then
      res = Log(es) / 1.81
    Else
      'Borenstein et. al (2009, p. 3)
      res = Log(es) * Sqr(3) / WorksheetFunction.Pi()
    End If
 
  'Odds Ratio target Yule Q
  ElseIf (from = "or" And target = "yuleq") Then
    res = (es - 1) / (es + 1)
  
  'Odds Ratio target Yule Y
  ElseIf (from = "or" And target = "yuley") Then
    res = (Sqr(es) - 1) / (Sqr(es) + 1)
 

  'OMEGA SQUARED
  'Omega squared target Epsilon squared
  ElseIf (from = "omegasq" And target = "epsilonsq") Then
    res = es / (1 - ex1 / (ex2 + ex1))


  'RANK BISERIAL
  'Rank Biserial target Vargha and Delaney A
  ElseIf (from = "rb" And target = "vda") Then
    res = (es + 1) / 2
    
    
  'VARGHA AND DELANEY A
  'Vargha and Delaney A target Rank Biserial
  ElseIf (from = "vda" And target = "rb") Then
    res = 2 * es - 1


  'YULE Q
  'Yule Q target Odds Ratio
  ElseIf (from = "yuleq" And target = "or") Then
    res = (1 + es) / (1 - es)

  'Yule Q target Yule Y
  ElseIf (from = "yuleq" And target = "yuley") Then
    res = (1 - Sqr(1 - es ^ 2)) / es


  'YULE Y
  'Yule Y target Odds Ratio
  ElseIf (from = "yuley" And target = "or") Then
    res = ((1 + es) / (1 - es)) ^ 2

  'Yule Y target Yule Q
  ElseIf (from = "yuley" And target = "yuleq") Then
    res = (2 * es) / (1 + es ^ 2)
  End If


  es_convert = res

End Function
