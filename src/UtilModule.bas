Attribute VB_Name = "UtilModule"
' ---
' screen, calculate Update
' ---
' �����X�V��~
Public Sub stopCalculate()
  Application.ScreenUpdating    = False
  ActiveSheet.EnableCalculation = False
  Application.Calculation       = xlCalculationManual
End Sub

' �����X�V�L��
Public Sub startCalculate()
  Application.Calculation       = xlCalculationAutomatic
  ActiveSheet.EnableCalculation = True
  Application.ScreenUpdating    = True
End Sub


' ---
' workbooks
' ---
' �ق��̃u�b�N���J���Ă���ꍇ���ׂĂ���鏈���B
Public Sub closeAllBooks()
  Do While Workbooks.Count >= 2
    For Each wb In Workbooks
      If wb.name <> ThisWorkbook.name Then
        Application.DisplayAlerts = Flase
        wb.Close saveChanges:=False
        Application.DisplayAlerts = True
      End If
    Next wb
  Loop
End Sub

' ---
' worksheet
' ---
' hasSheet
Public Function hasSheet(book, ByVal name As String)
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

' ---
' Cells
' ---
' current cell address
' Return an address of current cell.
Public Function curtAddr()
  curtAddr = Evaluate("ADDRESS(ROW(), COLUMN())")
End Function

' Get next number.
' Int cid
Public Function nextId(cid)
  Set c = Range(cid).Offset(-1)

  If Len(c) <= 0 Or Not IsNumeric(c) Then
    c = Range(cid).Offset(-2)
  End If

  c = c + 1
  nextId = c

  Set c = Nothing
End Function


' ---
' last row, col
' Worksheets o
' ---
' lastRow
Public Function lastRow(o, Optional first As Variant = 1)
  lastRow = o.Cells(Rows.Count, first).End(xlUp).Row
End Function

' lastColoumn
Public Function lastCol(o, Optional first As Variant = 1)
  lastCol = o.Cells(first, Columns.Count).End(xlToLeft).Column
End Function


' ---
' Sharpes
' ---
' set all sharpes position to be a fixed sharp.
Sub sharpFreeFloat()
  For Each ws In ThisWorkbook.Worksheets
    If ws.Shapes.Count > 0 Then
      For Each s In ws.Shapes
        If Not s.AutoShapeType = msoShapeMixed Then
          s.Placement = xlFreeFloating
        End If
      Next
    End If
  Next
End Sub


' ---
' echo message
' ---
' show message
Public Sub pMsg(msg, sec)
  Dim o As Object
  Set o = CreateObject("WScript.Shell")
  o.Popup msg, sec, "�����\��", vbInformation
  Set o = Nothing
End Sub


' ---
' html TAG
' ---
' tag
Public Function tag(tName As String, str As String)
  Set doc = New MSHTML.HTMLDocument
  Set t   = doc.createElement(tName)

  t.innerText = str
  tag         = t.outerHTML

  Set t   = Nothing
  Set doc = Nothing
End Function

' br
Public Function br()
  Set doc = New MSHTML.HTMLDocument
  Set t   = doc.createElement("br")

  br = t.outerHTML

  Set t   = Nothing
  Set doc = Nothing
End Function


' ---
' glob
' ---
Public Sub glob(fPath, ary)
  Dim fso As New Scripting.FileSystemObject

  For Each f In fso.GetFolder(fPath).files
    ary.Add f
  Next

  If fso.GetFolder(fPath).SubFolders.Count > 0 Then
    For Each d In fso.GetFolder(fPath).SubFolders
      ary.Add d
      glob d, ary
    Next
  End If

  Set fso = Nothing
End Sub


' ---
' array
' ---
' uniq
Function uniq(ary) As Object
  Set nAry = CreateObject("System.Collections.ArrayList")

  For Each v In ary
    If Not nAry.contains(v) Then
      nAry.Add v
    End If
  Next

  Set uniq = nAry
End Function


' ---
' color
' ---

' getRGB
' String c
'   ex: "255,255,255"
Function getRGB(c)
  myColor = Split(c, ",")
  getRGB = RGB(myColor(0), myColor(1), myColor(2))
End Function

' get interior color by Formula
Public Function getInteriorColorByFormula(c)
  Set fc = c.FormatConditions
  fLen = fc.Count
  cColor = 0

  For i = 1 To fLen
    If Evaluate(fc(i).Formula1) = c.Formula Then
      cColor = fc(i).Interior.color
    End If
  Next i

  getInteriorColorByFormula = cColor
  Set fc = Nothing
End Function


' ---
' Time
' ---
' String to (time)object
Public Function timeObject(Optional val = "00:00:00") As Object
  Set o = CreateObject("Scripting.Dictionary")
  tmp = Split(val, ":")
  o.Add "h", tmp(0)
  o.Add "m", tmp(1)
  o.Add "s", tmp(2)
  Set timeObject = o
  Set o = Nothing
End Function

' Convert time value to second. (Long)
Public Function time2sec(h, m, s)
  hourPerSec = 60 * 60
  minPerSec  = 60
  time2sec   = (h * hourPerSec) + (m * minPerSec) + s
End Function

' Convert string time value to second.
Public Function strtime2sec(strtime)
  Set t       = timeObject(strtime)
  strtime2sec = time2sec(t("h"), t("m"), t("s"))
  Set t       = Nothing
End Function

' Convert second to time string.
Public Function sec2time(s) As String
  sec2time = Application.WorksheetFunction.Text(CDate(s / 86400#), "[h]:mm:ss")
End Function

' Return the first day in set month.
Public Function getFirstDayInMonth(today)
  getFirstDayInMonth = DateSerial(Year(today), Month(today), 1)
End Function

' Return the last day in set month.
Public Function getLastDayInMonth(today)
  getLastDayInMonth = DateSerial(Year(today), Month(today) + 1, 0)
End Function

' Return the workday in month
Public Function getWorkdayInMonth(firstDay, lastDay, holidayRange)
  getWorkdayInMonth = Application.WorksheetFunction.NetworkDays(firstDay, lastDay, holidayRange)
End Function


' ---
' String
' ---
' regReplace
Public Function regReplace(val, rval, pat)
  Set re = CreateObject("VBScript.RegExp")
  With re
    .Pattern = pat
    .IgnoreCase = True
    .Global = True
  End With

  regReplace = re.Replace(val, rval)

  Set re = Nothing
End Function
