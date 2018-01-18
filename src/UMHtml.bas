Attribute VB_Name = "UMHtml"
' tag
Public Function tag(ByVal tName As String, ByVal label As String, Optional ByVal styleObj As Object = Nothing) As String
  Dim doc As MSHTML.HTMLDocument
  Dim t As MSHTML.HTMLElementCollection

  Set doc = New MSHTML.HTMLDocument
  Set t = doc.createElement(tName)
  t.innerText = label

  If Not styleObj Is Nothing Then
    Dim s As Variant
    For Each s In styleObj
      t.style.setAttribute s, styleObj(s)
    Next
  End If

  tag = t.outerHTML
  Set t = Nothing
  Set doc = Nothing
End Function



' br
Public Function br() As String
  Dim doc As MSHTML.HTMLDocument
  Dim t As MSHTML.HTMLElementCollection

  Set doc = New MSHTML.HTMLDocument
  Set t = doc.createElement("br")

  br = LCase(t.outerHTML)

  Set t = Nothing
  Set doc = Nothing
End Function

