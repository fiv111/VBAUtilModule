Attribute VB_Name = "UMRange"
' current cell address
' Return an address of current cell.
Public Function curtAddr() As String
  curtAddr = Evaluate("ADDRESS(ROW(), COLUMN())")
End Function



' Get next number.
' Int cid
Public Function nextId(addr As String) As Double
  Dim c As Range
  Dim val As Variant

  Set c = ActiveSheet.Range(addr).Offset(-1)
  val = c.Value

  If Len(val) <= 0 Or Not IsNumeric(val) Then
    val = nextId(c.Address)
    GoTo done
  End If

  val = val + 1

done:
  nextId = val
  Set c = Nothing
End Function



' ---
' last row, column
' Worksheets o
' ---
' lastRow
Public Function lastRow(ByVal o As Worksheet, Optional ByVal first As Integer = 1) As Long
  lastRow = o.Cells(Rows.Count, first).End(xlUp).Row
End Function



' lastColoumn
Public Function lastCol(ByVal o As Worksheet, Optional ByVal first As Integer = 1) As Long
  lastCol = o.Cells(first, Columns.Count).End(xlToLeft).Column
End Function
