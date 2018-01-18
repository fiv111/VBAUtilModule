Attribute VB_Name = "UMDigest"
Public Enum Digest
  SHA1
  SHA256
  SHA384
  SHA512
  HMACMD5
  HMACSHA1
  HMACSHA256
  HMACSHA384
  HMACSHA512
End Enum




' digestHash
Public Function digestHash(ByVal hashType As Integer, Optional ByVal secretKey As String = "", Optional ByVal str As String = "0", Optional ByVal lowercase = True) As String
  Dim digestObj, utf8Obj As Object
  Dim bytes() As Byte
  Dim hashList() As Byte
  Dim i, hash As Variant

  Set utf8Obj = CreateObject("System.Text.UTF8Encoding")
  Select Case hashType
  Case 0
    Set digestObj = CreateObject("System.Security.Cryptography.SHA1Managed")
  Case 1
    Set digestObj = CreateObject("System.Security.Cryptography.SHA256Managed")
  Case 2
    Set digestObj = CreateObject("System.Security.Cryptography.SHA384Managed")
  Case 3
    Set digestObj = CreateObject("System.Security.Cryptography.SHA512Managed")
  Case 4
    Set digestObj = CreateObject("System.Security.Cryptography.HMACMD5")
  Case 5
    Set digestObj = CreateObject("System.Security.Cryptography.HMACSHA1")
  Case 6
    Set digestObj = CreateObject("System.Security.Cryptography.HMACSHA256")
  Case 7
    Set digestObj = CreateObject("System.Security.Cryptography.HMACSHA384")
  Case 8
    Set digestObj = CreateObject("System.Security.Cryptography.HMACSHA512")
  Case Else
    Set digestObj = CreateObject("System.Security.Cryptography.SHA1Managed")
  End Select
  
  bytes = utf8Obj.GetBytes_4(secretKey & str)
  hashList = digestObj.ComputeHash_2(bytes)
  
  For i = 1 To UBound(hashList) + 1
    hash = hash & Right("0" & Hex(AscB(MidB(hashList, i, 1))), 2)
  Next i
  
  If lowercase Then
    digestHash = LCase(hash)
  Else
    digestHash = hash
  End If
End Function
