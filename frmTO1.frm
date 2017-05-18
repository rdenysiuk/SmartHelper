VERSION 5.00
Begin VB.Form frmTO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Технічне обслуговування комп'ютерної техніки"
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
   Icon            =   "frmTO1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   6300
   Begin VB.ComboBox cmb_work 
      Height          =   330
      Left            =   1800
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txt_CountLetters 
      Height          =   330
      Left            =   1800
      TabIndex        =   4
      Top             =   1200
      Width           =   2295
   End
   Begin VB.ComboBox cmb_type 
      Height          =   330
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.ComboBox cmb_dev 
      Height          =   330
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Вид роботи"
      Height          =   210
      Left            =   600
      TabIndex        =   7
      Top             =   1680
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Кількість аркушів"
      Height          =   210
      Left            =   240
      TabIndex        =   5
      Top             =   1200
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
      Caption         =   "Пристрій"
      Height          =   210
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   750
   End
End
Attribute VB_Name = "frmTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public myFrmMain As frmMain


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
    txt_CountLetters.Text = ""
End If

End With


ConnectToDataBase
myRS.Open SelInv, myADO, adOpenStatic

With cmb_type
Do While Not myRS.EOF
    .AddItem myRS("di_model").Value & " | Каб.-" & myRS("di_kab").Value
    .ItemData(.NewIndex) = myRS("di_nom").Value
    myRS.MoveNext
Loop
End With

err1:
If Err.Number = 381 Then cmb_type.Text = "(Оберіть пристрій)"

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
myFrmMain.teh.Enabled = True
End Sub

