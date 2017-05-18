Attribute VB_Name = "Exists"
Option Explicit

Public Function Folder_Exists(strPathName As String) As Boolean
Dim strDir As String
strDir = Dir(strPathName, vbDirectory)
If Len(strDir) = 0 Then
    Folder_Exists = False
Else
    Folder_Exists = True
End If
End Function

Public Function File_Exists(strPathFile As String) As Boolean
Dim strFile As String
strFile = Dir(strPathFile, vbNormal)
If Len(strFile) = 0 Then
    File_Exists = False
Else
    File_Exists = True
End If
End Function
