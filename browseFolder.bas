Attribute VB_Name = "Module2"
Option Explicit

Private Type BROWSEINFO
   hOwner As Long
   pidlRoot As Long
   pszDisplayName As String
   lpszTitle As String
   ulFlags As Long
   lpfn As Long
   lParam As Long
   lImage As Long
End Type

Private Type SHITEMID
   cb As Long
   abID As Byte
End Type

Private Type ITEMIDLIST
   mkid As SHITEMID
End Type

Private Const CSIDL_DESKTOP = &H0
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Public Function BrowseForFolder(lnghWnd As Long, strMessage As String, Optional DefaultPath As String) As String
   Dim biFolder As BROWSEINFO
   Dim idlList As ITEMIDLIST
   Dim lngIDLptr As Long
   Dim lngResult As Long
   Dim strPath As String

   On Error GoTo PROC_ERR
   SHGetSpecialFolderLocation lnghWnd, CSIDL_DESKTOP, idlList
   With biFolder
      .hOwner = lnghWnd
      .pidlRoot = idlList.mkid.cb
      .lpszTitle = strMessage
      .ulFlags = BIF_RETURNONLYFSDIRS
   End With
   lngIDLptr = SHBrowseForFolder(biFolder)
   strPath = Space$(260)
   lngResult = SHGetPathFromIDList(ByVal lngIDLptr, ByVal strPath)
   If lngResult <> 0 Then
      strPath = Left$(strPath, InStr(1, strPath, vbNullChar) - 1)
   Else
      strPath = DefaultPath
   End If
   BrowseForFolder = strPath

PROC_EXIT:
   Exit Function

PROC_ERR:
   MsgBox "Error: " & Err.Number & ". " & Err.Description
   Resume PROC_EXIT
End Function
