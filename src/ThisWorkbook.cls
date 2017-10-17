VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
  Call UtilModule.shapesFreeFloating
  
  Dim s As Variant
  Dim sh As Worksheet
  Set sh = ThisWorkbook.ActiveSheet
  
  For Each s In ThisWorkbook.Worksheets
    Call UtilModule.stopCalculate
    s.Activate
    s.Calculate
  Next
  
  sh.Activate
  Call UtilModule.startCalculate
  Set sh = Nothing
End Sub


Private Sub Workbook_SheetDeactivate(ByVal sh As Object)
  Call UtilModule.shapesFreeFloating
End Sub
