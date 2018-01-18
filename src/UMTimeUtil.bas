Attribute VB_Name = "UMTimeUtil"


' String to (time)object
Public Function timeObject(Optional ByVal val As String = "00:00:00") As Object
  Dim o As Object
  Dim tmp() As String

  Set o = CreateObject("Scripting.Dictionary")
  tmp = Split(val, ":")

  o.Add "h", CDbl(tmp(0))
  o.Add "m", CDbl(tmp(1))
  o.Add "s", CDbl(tmp(2))

  Set timeObject = o
  Set o = Nothing
End Function



' Convert time value to second. (Long)
Public Function time2sec(ByVal h As Double, ByVal m As Double, ByVal s As Double) As Double
  Dim hourPerSec As Double
  Dim minPerSec As Double
  hourPerSec = 60 * 60
  minPerSec = 60
  time2sec = (h * hourPerSec) + (m * minPerSec) + s
End Function



' Convert string time value to second.
Public Function strtime2sec(ByVal strtime As String) As Double
  Dim t As Object
  Set t = timeObject(strtime)
  strtime2sec = time2sec(t("h"), t("m"), t("s"))
  Set t = Nothing
End Function



' Convert second to time string.
Public Function sec2time(ByVal s As Double) As String
  sec2time = Application.WorksheetFunction.Text(CDate(s / 86400#), "[h]:mm:ss")
End Function



' Return the first day in set month.
Public Function getFirstDayInMonth(ByVal today As Variant) As Variant
  getFirstDayInMonth = DateSerial(Year(today), Month(today), 1)
End Function



' Return the last day in set month.
Public Function getLastDayInMonth(ByVal today As Variant) As Variant
  getLastDayInMonth = DateSerial(Year(today), Month(today) + 1, 0)
End Function



' Return the workday in month
Public Function getWorkdayInMonth(ByVal firstDay As Variant, ByVal lastDay As Variant, ByVal holidayRange As Variant) As Double
  getWorkdayInMonth = Application.WorksheetFunction.NetworkDays(firstDay, lastDay, holidayRange)
End Function
