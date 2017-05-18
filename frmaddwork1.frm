VERSION 5.00
Begin VB.Form frmAddWork 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Додати відомості про обслуговування"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6300
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmaddwork1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   6300
   Begin VB.TextBox txt_roz 
      Height          =   330
      Left            =   4200
      TabIndex        =   13
      Text            =   "0"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txt_prim 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      MaxLength       =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   2280
      Width           =   3735
   End
   Begin VB.CommandButton cmbAddWork 
      Caption         =   "Додати роботу"
      Height          =   450
      Left            =   3840
      TabIndex        =   8
      Top             =   3480
      Width           =   1695
   End
   Begin VB.ComboBox cmb_work 
      Height          =   330
      Left            =   1800
      TabIndex        =   6
      Top             =   1800
      Width           =   3735
   End
   Begin VB.TextBox txt_CountLetters 
      Height          =   330
      Left            =   1800
      TabIndex        =   4
      Text            =   "0"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ComboBox cmb_type 
      Height          =   330
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   3735
   End
   Begin VB.ComboBox cmb_dev 
      Height          =   330
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "* - обов'язкові поля для заповнення"
      ForeColor       =   &H80000015&
      Height          =   210
      Left            =   2400
      TabIndex        =   12
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Примітка"
      Height          =   210
      Left            =   840
      TabIndex        =   11
      Top             =   2280
      Width           =   750
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "пристрою *"
      Height          =   210
      Left            =   720
      TabIndex        =   9
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Вид роботи *"
      Height          =   210
      Left            =   600
      TabIndex        =   7
      Top             =   1800
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Кількість аркушів"
      Height          =   210
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1395
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Найменування"
      Height          =   210
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Тип пристрою *"
      Height          =   210
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1395
   End
End
Attribute VB_Name = "frmAddWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public myFrmMain As frmMain
Dim LastIdDev As Integer

Private Sub cmb_dev_Click()
cmb_type.Clear
End Sub

Private Sub cmb_dev_Validate(Cancel As Boolean)
Dim SelInv As String

On Error GoTo err1

With cmb_dev
SelInv = "SELECT di_nom, di_Model, di_kab FROM dev_inv WHERE Di_dev = " & .ItemData(.ListIndex) & " ORDER BY Di_model"

If .ItemData(.ListIndex) = 2 Then
    Label3.Enabled = True
    txt_CountLetters.Enabled = True
Else
    Label3.Enabled = False
    txt_CountLetters.Enabled = False
    txt_CountLetters.Text = "0"
End If

End With


ConnectToDataBase
myRS.Open SelInv, myADO, adOpenStatic

With cmb_type
Do While Not myRS.EOF
    .AddItem myRS("di_nom").Value & " | " & myRS("di_model").Value & " | Каб.-" & myRS("di_kab").Value
    .ItemData(.NewIndex) = myRS("di_nom").Value
    myRS.MoveNext
Loop
End With

err1:
If Err.Number = 381 Then cmb_type.Text = "(Оберіть пристрій)"

End Sub


Private Sub cmb_type_Validate(Cancel As Boolean)
Dim LastId As String
Dim I As Integer

If Len(cmb_type.Text) = 0 Or cmb_type.Text = "(Оберіть пристрій)" Then Exit Sub

LastId = "SELECT max(sto_date) as dat, sto_inv, sto_id FROM sto GROUP BY sto_inv, sto_id HAVING sto_inv=" & _
    cmb_type.ItemData(cmb_type.ListIndex) & " order by max(sto_date) desc"
ConnectToDataBase
myRS.Open LastId, myADO, adOpenStatic

If myRS.RecordCount > 1 Then
    For I = 1 To myRS.RecordCount
    If I = 1 Then LastIdDev = myRS("sto_id").Value
    Next I
ElseIf myRS.RecordCount = 1 Then
    LastIdDev = myRS("sto_id").Value
Else
    LastIdDev = 0
End If

myRS.close
End Sub

Private Sub cmbAddWork_Click()
Dim InsertWork As String
Dim sTime As String
On Error GoTo CmbAddWorkErr

If Len(cmb_dev.Text) = 0 Or Len(cmb_type.Text) = 0 Or Len(cmb_work.Text) = 0 Then
MsgBox "Обов'язкові поля не заповнені.", vbExclamation + vbOKOnly, "Увага"
Exit Sub
End If

With txt_prim
If Len(.Text) = 0 Then
    .Text = "NULL"
    .Enabled = False
Else
    .Text = "'" & .Text & "'"
End If
End With

If Len(Hour(Time)) = 1 Then
    sTime = "0" & Hour(Time) & ":" & Minute(Time)
Else
    sTime = Hour(Time) & ":" & Minute(Time)
End If

InsertWork = "INSERT INTO sto (sto_inv, sto_works, sto_countlet, sto_let, sto_date, sto_prog, sto_prim) " & _
    "VALUES (" & cmb_type.ItemData(cmb_type.ListIndex) & "," & _
    cmb_work.ItemData(cmb_work.ListIndex) & "," & _
    txt_CountLetters.Text & "," & txt_roz.Text & ",'" & _
    Date & " " & sTime & "'," & _
    ReadINI("viezd", "ID", PathFileIni) & "," & txt_prim & ")"
ConnectToDataBase

myRS.Open InsertWork, myADO, adOpenForwardOnly


'If MsgBox("Додано: " & cmb_work & vbTab & cmb_dev & vbTab & cmb_type & vbCrLf & vbCrLf & _
'    "Додачи ще відомості про обслуговування техніки?", vbExclamation + vbYesNo, "Додано роботу") = vbYes Then
'    cmb_dev.SetFocus
'    cmb_type.Clear
'    txt_CountLetters.Text = 0
'    txt_roz.Text = 0
'    txt_CountLetters.Enabled = False
'    cmb_work.Text = ""
'    txt_prim.Text = ""
'Else
'    Unload Me
'End If

cmb_dev.Text = ""
cmb_type.Text = ""
txt_CountLetters.Text = "0"
cmb_work.Text = ""
txt_prim.Text = ""
Exit Sub

CmbAddWorkErr:
MsgBox Err.Description, vbExclamation, "Warning " & Err.Number
Exit Sub

End Sub

Private Sub Form_Load()
Dim SelDevice As String
Dim SelWorks As String

Label3.Enabled = False
txt_CountLetters.Enabled = False

SelDevice = "SELECT * FROM device"

ConnectToDataBase

myRS.Open SelDevice, myADO, adOpenStatic

With cmb_dev
Do While Not myRS.EOF
    .AddItem myRS("dev_name").Value
    .ItemData(.NewIndex) = myRS("dev_id").Value
    myRS.MoveNext
Loop
End With

myRS.close

SelWorks = "SELECT * FROM works ORDER BY w_name"
myRS.Open SelWorks, myADO, adOpenDynamic

With cmb_work
Do While Not myRS.EOF
    .AddItem myRS("w_name").Value
    .ItemData(.NewIndex) = myRS("w_id").Value
    myRS.MoveNext
Loop
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
myFrmMain.addWork.Enabled = True
End Sub

Private Sub txt_CountLetters_Validate(Cancel As Boolean)
Dim sLastLetter As String, iLastLetter As Long
sLastLetter = "SELECT STO_CountLet FROM sto WHERE sto_id = " & LastIdDev

ConnectToDataBase
myRS.Open sLastLetter, myADO, adOpenStatic

If myRS.RecordCount = 1 Then
    iLastLetter = Val(myRS("STO_CountLet").Value)
    txt_roz.Text = Val(txt_CountLetters.Text) - iLastLetter
Else
    iLastLetter = 0
End If

End Sub

