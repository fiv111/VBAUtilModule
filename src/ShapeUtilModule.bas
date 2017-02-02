Attribute VB_Name = "ShapeUtilModule"
Option Explicit

' ---
' Change All the shape object placement to be free-floating in those worksheets.
' ---
Public Sub shapesFreeFloating()
  Dim ws, s As Variant

  For Each ws In ThisWorkbook.Worksheets
    If ws.Shapes.Count > 0 Then
      For Each s In ws.Shapes
        If Not s.AutoShapeType = msoShapeMixed Or Not s.Type = msoShapeMixed Then
          s.Placement = xlFreeFloating
        End If
      Next
    End If
  Next
End Sub


' ---
' Get the information of currentsheet from the other worksheet.
' ---
Public Function getCurrentSheetInfo(ByVal sheetName As String, ByVal r As String, Optional ByVal num As Long = 0) As Variant
  Dim val As Variant

  If num = 0 Then
    val = ThisWorkbook.Worksheets(sheetName).Range(r & CInt(ThisWorkbook.ActiveSheet.Name) + 1).Value
  Else
    val = ThisWorkbook.Worksheets(sheetName).Range(r & num + 1).Value
  End If

  If Len(val) <= 0 Then
    getCurrentSheetInfo = "-"
  Else
    getCurrentSheetInfo = val
  End If
End Function


' ---
' Select all shape object from current worksheet.
' ---
Public Sub selectCurrentWorksheetShapes()
  If ActiveSheet.Shapes.Count > 0 Then
    ActiveSheet.Shapes.SelectAll
  End If
End Sub


' ---
' Attach a numbering label for selected shape.
' ---
' attach Label
Public Sub attachNumberingLabel()
  Dim sWidth, sHeight As Double
  Dim i As Variant
  Dim s, sp As Shape
  Dim labelName, fontName As String
  Dim bgColor, borderColor As Long

  ' Label Shape size. 1cm
  sWidth = 28.2
  sHeight = 28.2

  fontName = "メイリオ"
  labelName = "VBAWFLabel"

  ' yellow bg
  ' black line
  bgColor = RGB(255, 255, 0)
  borderColor = RGB(0, 0, 0)

  If Not TypeName(Selection) = "Range" Then

    For i = 1 To Selection.ShapeRange.Count
      Set sp = Selection.ShapeRange.Item(i)
      Set s = ActiveSheet.Shapes.AddShape(msoShapeRectangle, sp.Left + sp.width - sWidth * 0.5, sp.Top - sHeight * 0.5, sWidth, sHeight)

      With s
        .Name = labelName & i
        .Fill.ForeColor.RGB = bgColor
        .Line.ForeColor.RGB = borderColor
        .Line.Weight = 3
        .TextFrame.Characters.Text = i
        .TextFrame.Characters.Font.Color = borderColor
        .TextFrame.Characters.Font.Size = 10
        .TextFrame.Characters.Font.Name = fontName
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
      End With

      Set s = Nothing
      Set sp = Nothing
    Next
  End If
End Sub


' ---
' Attach numbering labels for all shapes on current worksheet.
' ---
Public Sub attachNumberingLabelAll()
  Call selectCurrentWorksheetShapes
  Call attachNumberingLabel
End Sub


' ---
' Refresh
' ---
Public Sub refresh()
  Application.ScreenUpdating = False
  ActiveSheet.EnableCalculation = False

  Application.Calculation = xlCalculationManual
  ThisWorkbook.ActiveSheet.Calculate
  Application.Calculation = xlCalculationAutomatic

  ActiveSheet.EnableCalculation = True
  Application.ScreenUpdating = True
End Sub
