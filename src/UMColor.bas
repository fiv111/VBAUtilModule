Attribute VB_Name = "UMColor"
' getRGB
' String c
'   ex: "255,255,255"
Public Function getRGB(Optional ByVal c As String = "0,0,0") As Long
  Dim myColor() As String
  myColor = Split(c, ",")
  getRGB = RGB(myColor(0), myColor(1), myColor(2))
End Function



' get interior color by Formula
Public Function getInteriorColorByFormula(ByVal c As Range) As Long
  Dim fc As FormatConditions
  Dim fLen As Long
  Dim cColor As Long
  Dim i As Integer

  Set fc = c.FormatConditions
  fLen = fc.Count
  cColor = 0

  For i = 1 To fLen
    If Evaluate(fc(i).Formula1) = c.Formula Then
      cColor = fc(i).Interior.Color
      Exit For
    End If
  Next i

  getInteriorColorByFormula = cColor
  Set fc = Nothing
End Function

