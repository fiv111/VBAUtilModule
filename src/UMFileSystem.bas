Attribute VB_Name = "UMFileSystem"
Public Sub glob(ByVal fPath As String, ByRef ary As Object)
  Dim fso As New Scripting.FileSystemObject

  Dim f As Variant
  For Each f In fso.GetFolder(fPath).files
    ary.Add f
  Next

  If fso.GetFolder(fPath).SubFolders.Count > 0 Then
    Dim d As Variant
    For Each d In fso.GetFolder(fPath).SubFolders
      ary.Add d
      glob d, ary
    Next
  End If

  Set fso = Nothing
End Sub
