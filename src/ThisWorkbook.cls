VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
  Call UMShapes.shapesFreeFloating
  
  Dim s As Variant
  Dim sh As Worksheet
  Set sh = ThisWorkbook.ActiveSheet
  
  For Each s In ThisWorkbook.Worksheets
    Call UMFunc.stopCalculate
    s.Activate
    s.Calculate
  Next
  
  sh.Activate
  Call UMFunc.startCalculate
  Set sh = Nothing
End Sub


Private Sub Workbook_SheetDeactivate(ByVal sh As Object)
  Call UMShapes.shapesFreeFloating
End Sub
