VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutoCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'add to class module http://www.tek-tips.com/viewthread.cfm?qid=757158
Option Explicit
Const CB_GETEXTENDEDUI = &H156
Const CB_SETEXTENDEDUI = &H155
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
  (ByVal hwnd As Long, ByVal wMsg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) _
  As Long
Const CB_FINDSTRING = &H14C
Const CB_ERR = (-1)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
  (ByVal hwnd As Long, ByVal wMsg As Long, _
   ByVal wParam As Long, lParam As Any) _
  As Long
Private Declare Function MoveWindow Lib "user32" _
       (ByVal hwnd As Long, ByVal x As Long, ByVal y As _
       Long, ByVal nWidth As Long, ByVal nHeight As Long, _
       ByVal bRepaint As Long) As Long

Public Sub SetHeight(ByRef cboThis As ComboBox, CHeight As Long)
   Dim ret
   Dim iScaleMode
   iScaleMode = cboThis.Parent.ScaleMode
   cboThis.Parent.ScaleMode = vbPixels
   ret = MoveWindow(cboThis.hwnd, cboThis.Left, cboThis.Top, cboThis.Width, CHeight, 1)
   cboThis.Parent.ScaleMode = iScaleMode
End Sub

Public Property Let ComboExtendedUI(ByRef cboThis As ComboBox, ByVal bState As Boolean)
   ' Set whether combo box drops down using the Down Arrow or not:
   SendMessageLong cboThis.hwnd, CB_SETEXTENDEDUI, Abs(bState), 0
End Property

Public Property Get ComboExtendedUI(ByRef cboThis As ComboBox) As Boolean
   ' Get whether combo box drops down using the Down Arrow or not:
   ComboExtendedUI = (SendMessageLong(cboThis.hwnd, CB_GETEXTENDEDUI, 0, 0) <> 0)
End Property

Public Function AutoFind(ByRef cboCurrent As ComboBox, _
                        ByVal KeyAscii As Integer, _
                        Optional ByVal LimitToList As Boolean = False)
       
Dim lCB As Long
Dim sFindString As String

On Error GoTo Err_Handler
   If KeyAscii = 8 Then
       If cboCurrent.SelStart <= 1 Then
           cboCurrent = ""
           AutoFind = 0
           Exit Function
       End If
       If cboCurrent.SelLength = 0 Then
           sFindString = UCase(Left(cboCurrent, Len(cboCurrent) - 1))
       Else
           sFindString = Left$(cboCurrent.Text, cboCurrent.SelStart - 1)
       End If
   'ElseIf KeyAscii < 32 Or KeyAscii > 255 Then
       'Exit Function
       '=====================
    ElseIf KeyAscii = 13 Then
       SendKeys "{Tab}"
    ElseIf KeyAscii < 32 Or KeyAscii > 255 Then
       Exit Function
       '=====================
   Else
       If cboCurrent.SelLength = 0 Then
           sFindString = UCase(cboCurrent.Text & Chr$(KeyAscii))
       Else
           sFindString = Left$(cboCurrent.Text, cboCurrent.SelStart) & Chr$(KeyAscii)
       End If
   End If
   lCB = SendMessage(cboCurrent.hwnd, CB_FINDSTRING, -1, ByVal sFindString)

   If lCB <> CB_ERR Then
       cboCurrent.ListIndex = lCB
       cboCurrent.SelStart = Len(sFindString)
       cboCurrent.SelLength = Len(cboCurrent.Text) - cboCurrent.SelStart
       AutoFind = 0
   Else
       If LimitToList = True Then
           AutoFind = 0
       Else
           AutoFind = KeyAscii
       End If
   End If
    
Err_Handler:

End Function



