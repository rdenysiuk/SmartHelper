Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function apiCharToOemBuff Lib "user32" Alias "CharToOemBuffA" (ByVal lpszSrc As String, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
Private Declare Function apiOemToCharBuff Lib "user32" Alias "OemToCharBuffA" (ByVal lpszSrc As String, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long

Public Function Win2Dos(pString As Variant) As String
Dim strBuffer As String, lngLen As Long
Win2Dos = "" & pString: lngLen = Len(Win2Dos) + 1
If lngLen = 1 Then Exit Function
strBuffer = String(lngLen, Chr(0))
apiCharToOemBuff Win2Dos, strBuffer, lngLen
Win2Dos = Left$(strBuffer, lngLen - 1)
End Function

Public Function Dos2Win(pString As Variant) As String
Dim strBuffer As String, lngLen As Long
Dos2Win = "" & pString: lngLen = Len(Dos2Win) + 1
If lngLen = 1 Then Exit Function
strBuffer = String(lngLen, Chr(0))
apiOemToCharBuff Dos2Win, strBuffer, lngLen
Dos2Win = Left$(strBuffer, lngLen - 1)
End Function
