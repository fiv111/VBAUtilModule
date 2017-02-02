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

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
  If Success Then
    Call ShapeUtilModule.refresh
  End If
End Sub


Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
  Call ShapeUtilModule.shapesFreeFloating
End Sub


Private Sub Workbook_Open()
  Call ShapeUtilModule.shapesFreeFloating
End Sub


Private Sub Workbook_SheetActivate(ByVal Sh As Object)
  Call ShapeUtilModule.refresh
End Sub
