VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNDI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Нормативно довідкова інформація"
   ClientHeight    =   3990
   ClientLeft      =   240
   ClientTop       =   435
   ClientWidth     =   5385
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNDI1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5385
   Begin VB.TextBox txt_NdiId 
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmb_Add 
      Caption         =   "Зберегти"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   2400
      Width           =   2415
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   330
      Left            =   3720
      TabIndex        =   4
      Top             =   1560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txt_kol 
      Height          =   325
      Left            =   2400
      TabIndex        =   2
      Text            =   "0"
      Top             =   1560
      Width           =   1575
   End
   Begin MSMask.MaskEdBox txt_date 
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##.##.####"
      PromptChar      =   "_"
   End
   Begin VB.Label Lbl_Status 
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Кі-сть введених змін"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Дата виконання"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "frmNDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public myFrmMain As frmMain
Dim PriznInsert As Boolean

Private Sub cmb_Add_Click()
Dim sqlUPD, sqlINS
sqlUPD = "UPDATE ndi SET ndi_kol= " & txt_kol.Text & " WHERE ndi_id = " & txt_NdiId.Text
sqlINS = "INSERT INTO ndi (ndi_date, ndi_prog, ndi_kol) " & _
    "VALUES ('" & txt_date & "'," & ReadINI("viezd", "ID", PathFileIni) & "," & txt_kol & ");"
ConnectToDataBase
If txt_kol.Text > 0 Then
    If PriznInsert = True Then
    myRS.Open sqlINS, myADO, adOpenStatic
    MsgBox "Дані додані", vbInformation, ":-)"
    Unload Me
    Else
    myRS.Open sqlUPD, myADO, adOpenStatic
    MsgBox "Дані оновлено", vbInformation, ":-)"
    Unload Me
    End If
Else
    MsgBox "Кількість змін не може дорівнювати 0", vbCritical, "Atention "
End If
End Sub


Private Sub Form_Load()
Dim selPresent
myFrmMain.ndi.Enabled = False
txt_date.Text = Date
selPresent = "SELECT * FROM ndi WHERE ndi_date = '" & txt_date & "' and ndi_prog = " & ReadINI("viezd", "ID", PathFileIni)
ConnectToDataBase
myRS.Open selPresent, myADO, adOpenStatic

If Val(myRS.RecordCount) > 0 Then
    txt_kol.Text = myRS("ndi_kol").Value
    txt_NdiId.Text = myRS("ndi_ID").Value
    StatusRec (True)
    PriznInsert = False
Else
    StatusRec (False)
    PriznInsert = True
End If
'MsgBox myRS.RecordCount
With Me
.Width = 5500
.Height = 4500
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
myFrmMain.ndi.Enabled = True
End Sub

Private Sub txt_date_KeyPress(KeyAscii As Integer)
TextPressEnter (KeyAscii)
End Sub


Private Sub txt_date_Validate(Cancel As Boolean)
Dim LookForRec As String
LookForRec = "SELECT * FROM ndi WHERE ndi_date = '" & txt_date.Text & "' and ndi_prog = " & ReadINI("viezd", "ID", PathFileIni)

ConnectToDataBase
myRS.Open LookForRec, myADO, adOpenStatic
If myRS.RecordCount > 0 Then
    txt_kol.Text = myRS("ndi_kol").Value
    txt_NdiId.Text = myRS("ndi_id").Value
    StatusRec (True)
Else
    txt_kol.Text = 0
    txt_NdiId.Text = ""
    StatusRec (False)
End If

End Sub

Private Sub UpDown1_DownClick()
txt_kol.SetFocus
With txt_kol
If .Text <> 0 Then
    .Text = .Text - 1
End If
End With

End Sub

Private Sub UpDown1_UpClick()
txt_kol.SetFocus
With txt_kol
.Text = .Text + 1
End With
End Sub

Private Function StatusRec(przn As Boolean) 'As Boolean

If przn = True Then
    With Lbl_Status
    .Caption = "Дані присутні на встановлену дату"
    .FontBold = True
    .ForeColor = vbBlue
    End With
Else
    With Lbl_Status
    .Caption = "Немає даних на встановлену дату"
    .FontBold = True
    .ForeColor = vbRed
    End With
End If

End Function
