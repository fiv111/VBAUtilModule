Attribute VB_Name = "UMFunc"
Option Explicit

Public Const kFontName = "Meiryo"
Public Const kFontSize = 10

' size. 1cm
Public Const kWidth1cm = 28.2
Public Const kHeight1cm = 28.2



' ---
' screen, calculate Update
' ---
' stopCalculate
Public Sub stopCalculate()
  Application.ScreenUpdating = False
  ActiveSheet.EnableCalculation = False
  Application.Calculation = xlCalculationManual
End Sub



' startCalculate
Public Sub startCalculate()
  Application.Calculation = xlCalculationAutomatic
  ActiveSheet.EnableCalculation = True
  Application.ScreenUpdating = True
End Sub



' ---
' echo message
' ---
' show message
Public Sub echo(ByVal msg As String, Optional ByVal sec As Integer = 1)
  Dim o As Object
  Set o = CreateObject("WScript.Shell")
  o.Popup msg, sec, "Auto Display", vbInformation
  Set o = Nothing
End Sub
