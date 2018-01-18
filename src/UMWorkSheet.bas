Attribute VB_Name = "UMWorkSheet"
' hasSheet
Public Function hasSheet(ByVal book As Workbook, ByVal name As String) As Boolean
  Dim s As Variant
  For Each s In book.Worksheets
    If s.name = name Then
      hasSheet = True
      GoTo fin
    Else
      hasSheet = False
    End If
  Next
fin:
End Function



' Get the information of currentsheet from the other worksheet.
Public Function getCurrentSheetInfo(ByVal sheetName As String, ByVal r As String, Optional ByVal num As Long = 0) As Variant
  Dim val As Variant

  Dim re As Object
  Set re = CreateObject("VBScript.RegExp")
  With re
    .Pattern = "^[0-9]+$"
    .IgnoreCase = True
    .Global = True
  End With
  
  If re.test(ThisWorkbook.ActiveSheet.name) Then
    If num = 0 Then
      val = ThisWorkbook.Worksheets(sheetName).Range(r & CInt(ThisWorkbook.ActiveSheet.name) + 1).Value
    Else
      val = ThisWorkbook.Worksheets(sheetName).Range(r & num + 1).Value
    End If
  Else
    val = "-"
  End If

  If Len(val) <= 0 Then
    getCurrentSheetInfo = "-"
  Else
    getCurrentSheetInfo = val
  End If
End Function



' Create a attribute table.
Private Sub attachAttributeTable(Optional ByVal ac As Range = Nothing)
  ' selection
  Dim s As Range

  If ac Is Nothing Then
    Set s = ActiveCell
  Else
    Set s = ac
  End If

  If s.Count > 1 Then
    MsgBox ("More then one cell has been selected. Select only ONE cell. Try again.")
    Exit Sub
  End If

  ' head
  Dim docHead(5), attrHead(5), docVal(5) As String

  docHead(0) = "PageID"
  docHead(1) = "PageName"
  docHead(2) = "CreatedBy"
  docHead(3) = "UpdatedBy"
  docHead(4) = "CreatedAt"
  docHead(5) = "UpdatedAt"

  docVal(0) = "=getCurrentSheetInfo(" & Chr(34) & "Sitemap" & Chr(34) & ", " & Chr(34) & "A" & Chr(34) & ")"
  docVal(1) = "=getCurrentSheetInfo(" & Chr(34) & "Sitemap" & Chr(34) & ", " & Chr(34) & "B" & Chr(34) & ")"
  docVal(2) = "-"
  docVal(3) = "-"
  docVal(4) = Date
  docVal(5) = "=TODAY()"

  attrHead(0) = "ID"
  attrHead(1) = "Name"
  attrHead(2) = "Type"
  attrHead(3) = "Description"
  attrHead(4) = "Action"
  attrHead(5) = "Destination"

  ' color
  Dim docInfoHeadColor, attrHeadColor, colorWhite, colorGrey As Long
  docInfoHeadColor = RGB(51, 102, 153)
  attrHeadColor = RGB(128, 128, 128)
  colorWhite = RGB(255, 255, 255)
  colorGrey = RGB(80, 80, 80)

  Call UMFunc.stopCalculate

  ' print document head
  Dim i As Variant
  Dim r, c As Long
  For i = 0 To UBound(docHead)
    r = s.Row
    c = s.Column + i

    With ActiveSheet.Cells(r, c)
      .Value = docHead(i)
      .Interior.Color = docInfoHeadColor
      .Font.name = kFontName
      .Font.Color = colorWhite
      .Font.Bold = True
      .Font.size = kFontSize
      .Borders.Color = colorGrey
      .Borders.Weight = xlThin
      .Borders.LineStyle = xlContinuous
    End With

    With ActiveSheet.Cells(r + 1, c)
      .Value = docVal(i)
      .Font.name = kFontName
      .Borders.Color = colorGrey
      .Borders.Weight = xlThin
      .Borders.LineStyle = xlContinuous
    End With
  Next

  Dim j As Variant
  Dim attrSize As Integer
  attrSize = 10

  ' print attribute head
  For i = 0 To UBound(attrHead)
    r = s.Row + 2
    c = s.Column + i

    With ActiveSheet.Cells(r, c)
      .Value = attrHead(i)
      .Interior.Color = attrHeadColor
      .Font.name = kFontName
      .Font.Color = colorWhite
      .Font.Bold = True
      .Font.size = kFontSize
      .Borders.Color = colorGrey
      .Borders.Weight = xlThin
      .Borders.LineStyle = xlContinuous
    End With

    For j = 1 To attrSize
      If i = 0 Then
        ActiveSheet.Cells(r + j, c).Value = j
      Else
        ActiveSheet.Cells(r + j, c).Value = "-"
      End If

      With ActiveSheet.Cells(r + j, c)
        .Font.size = kFontSize
        .Font.name = kFontName
        .Borders.Color = colorGrey
        .Borders.Weight = xlThin
        .Borders.LineStyle = xlContinuous
      End With
    Next
  Next

  Call UMFunc.startCalculate
  Set s = Nothing
End Sub



' Create a attribute table for current worksheet.
Public Sub drawAttributeTable()
  attachAttributeTable ActiveCell
End Sub



' Create a attribute table for all worksheets.
Public Sub drawAttributeTableAll()
  Dim s As Range
  Set s = ActiveCell

  Dim shts As sheets
  Set shts = Worksheets

  Dim re As Object
  Set re = CreateObject("VBScript.RegExp")
  With re
    .Pattern = "^[0-9]+$"
    .IgnoreCase = True
    .Global = True
  End With
  
  Dim ws As Variant
  For Each ws In shts
    If re.test(ws.name) Then
      ws.Activate
      attachAttributeTable s
    End If
  Next

  Set shts = Nothing
  Set s = Nothing
End Sub


' Rename worksheets to ordered number.
Public Sub mvWorksheets()
  Dim num, shtSize As Long
  num = 1

  Dim ss As sheets
  Set ss = ActiveWindow.SelectedSheets
  shtSize = ss.Count

  Dim i As Variant
  For i = num To shtSize
    ss.Item(i).name = Math.Rnd() * Now
  Next

  For i = num To shtSize
    ss.Item(i).name = i
  Next

  Set ss = Nothing
End Sub

' show hidden worksheets
Sub showWorksheets()
  Dim ws As sheets
  Set ws = ThisWorkbook.Worksheets

  Dim s As Variant
  For Each s In ws
    If Not s.Visible Then
      s.Visible = True
    End If
  Next

  Set ws = Nothing
End Sub
