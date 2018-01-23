Attribute VB_Name = "UMShapes"
' Change All the shape object placement to be free-floating in those worksheets.
Public Sub shapesFreeFloating()
  Dim ws, s As Variant

  For Each ws In ThisWorkbook.Worksheets
    If ws.Shapes.Count > 0 Then
      For Each s In ws.Shapes
        If (s.Type = msoPicture) Or (s.Type = msoTextBox) Or _
           (s.Type = msoShapeTypeMixed) Or (s.Type = msoOLEControlObject) Or _
           (s.Type = msoLinkedPicture) Or (s.Type = msoLinkedOLEObject) Or _
           (s.Type = msoLine) Or (s.Type = msoGroup) Or _
           (s.Type = msoChart) Or (s.Type = msoCanvas) Or _
           (s.Type = msoAutoShape) Then
             s.Placement = xlFreeFloating
        End If
      Next
    End If
  Next
End Sub



' Select all shape object from current worksheet.
Public Sub selectCurrentWorksheetShapes()
  If ActiveSheet.Shapes.Count > 0 Then
    ActiveSheet.Shapes.SelectAll
  End If
End Sub



' Attach a numbering label for selected shape.
Public Sub attachNumberingLabel()
  Dim i As Variant
  Dim s, sp As Shape
  Dim labelName As String
  Dim bgColor, borderColor As Long

  labelName = "VBAWFLabel"

  ' yellow bg
  ' black line
  bgColor = RGB(255, 255, 0)
  borderColor = RGB(0, 0, 0)
  
  If Not TypeName(Selection) = "Range" Then
    For i = 1 To Selection.ShapeRange.Count
      Set sp = Selection.ShapeRange.Item(i)
      Set s = ActiveSheet.Shapes.AddShape(msoShapeRectangle, sp.Left + sp.Width - UMFunc.kWidth1cm * 0.5, sp.Top - UMFunc.kHeight1cm * 0.5, UMFunc.kWidth1cm, UMFunc.kHeight1cm)

      With s
        .name = labelName & i
        .Fill.ForeColor.RGB = bgColor
        .Line.ForeColor.RGB = borderColor
        .Line.Weight = 3
        .TextFrame.Characters.Text = i
        .TextFrame.Characters.Font.name = UMFunc.kFontName
        .TextFrame.Characters.Font.Color = borderColor
        .TextFrame.Characters.Font.size = UMFunc.kFontSize
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.AutoSize = True
      End With

      Set s = Nothing
      Set sp = Nothing
    Next
  End If
End Sub



' Attach numbering labels for all shapes on current worksheet.
Public Sub attachNumberingLabelAll()
  Call selectCurrentWorksheetShapes
  Call attachNumberingLabel
End Sub



' connect shapes
Public Sub shapesConnect()
  Dim arr, s1, s2 As Shape
  Dim ss As ShapeRange
  Dim i, size As Integer
  
  Set ss = Selection.ShapeRange
  size = ss.Count
  
  For i = 1 To size
    If i < size Then
      Set s1 = ss(i)
      Set s2 = ss(i + 1)
      
      Set arr = ActiveSheet.Shapes.AddConnector(msoConnectorElbow, s1.Left, s1.Top, s1.Left + s1.Width, s1.Top + s1.Height)
      arr.Line.EndArrowheadStyle = msoArrowheadTriangle
      arr.Line.Weight = 2
      arr.Line.ForeColor.RGB = RGB(0, 0, 0)
      
      With arr.ConnectorFormat
        .BeginConnect s1, 4
        .EndConnect s2, 4
      End With
      arr.RerouteConnections
      
      Set arr = Nothing
      Set s1 = Nothing
      Set s2 = Nothing
    End If
  Next
End Sub


' convert text to ShapeRectangle object
Public Sub txt2shapeRectangle()
  Dim r, selectedRange As Range
  Dim s As Shape
  Dim shapeName As String
  Dim sWidth, sHeight As Double
  Dim bgColor, borderColor As Long
      
  Set selectedRange = Selection
  
  shapeName = "VBAWFSitemapLabel"
  
  ' light blue bg
  ' black line
  bgColor = RGB(222, 235, 247)
  borderColor = RGB(0, 0, 0)
  
  UMFunc.stopCalculate
  
  For Each r In selectedRange
    If Len(r.Value) > 0 Then
      Set s = ActiveSheet.Shapes.AddShape(msoShapeRectangle, r.Left, r.Top, UMFunc.kWidth1cm, UMFunc.kHeight1cm)
      With s
        .name = shapeName & i
        .Fill.ForeColor.RGB = bgColor
        .Line.ForeColor.RGB = borderColor
        .Line.Weight = 1
        .TextFrame.Characters.Font.name = UMFunc.kFontName
        .TextFrame.Characters.Font.Color = borderColor
        .TextFrame.Characters.Font.size = UMFunc.kFontSize
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter

        .TextFrame.Characters.Text = r.Value
        .TextFrame.AutoSize = True
        .TextFrame.MarginLeft = UMFunc.kWidth1cm * 0.1
        .TextFrame.MarginRight = UMFunc.kWidth1cm * 0.1
        .TextFrame.MarginTop = UMFunc.kHeight1cm * 0.1
        .TextFrame.MarginBottom = UMFunc.kHeight1cm * 0.1
      End With
      
      Set s = Nothing
    End If
    r.Value = ""
  Next
  
  UMFunc.startCalculate
  
  Set selectedRange = Nothing
End Sub
