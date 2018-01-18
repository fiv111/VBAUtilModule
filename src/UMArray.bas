Attribute VB_Name = "UMArray"

' uniq
Public Function uniq(ByVal ary As Object) As Object
  Dim nAry As Object
  Set nAry = CreateObject("System.Collections.ArrayList")

  Dim v As Variant
  For Each v In ary
    If Not nAry.contains(v) Then
      nAry.Add v
    End If
  Next

  Set uniq = nAry
End Function
