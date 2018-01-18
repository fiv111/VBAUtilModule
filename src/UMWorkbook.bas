Attribute VB_Name = "UMWorkbook"


' close all workbooks
Public Sub closeAllBooks()
  Dim wb As Variant
  Do While Workbooks.Count >= 2
    For Each wb In Workbooks
      If wb.name <> ThisWorkbook.name Then
        Application.DisplayAlerts = False
        wb.Close saveChanges:=False
        Application.DisplayAlerts = True
      End If
    Next wb
  Loop
End Sub
