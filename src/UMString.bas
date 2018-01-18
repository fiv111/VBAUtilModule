Attribute VB_Name = "UMString"

' reReplace
Public Function reReplace(ByVal val As String, ByVal rval As String, ByVal pat As String) As String
  Dim re As Object
  Set re = CreateObject("VBScript.RegExp")

  With re
    .Pattern = pat
    .IgnoreCase = True
    .Global = True
  End With

  reReplace = re.Replace(val, rval)
  Set re = Nothing
End Function
