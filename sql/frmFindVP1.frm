VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFindVP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Пошук вигрузки 1LS"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8670
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Пошук"
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Район"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "N ОР"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ПІБ"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Дата"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Інспектор"
         Object.Width           =   1852
      EndProperty
   End
   Begin VB.TextBox txtTarget 
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      Top             =   960
      Width           =   3975
   End
   Begin VB.OptionButton pib 
      Caption         =   "За ПІБ пенсіонера"
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.OptionButton osob 
      Caption         =   "За номером ОР"
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblTarget 
      Caption         =   "Label1"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "frmFindVP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public myFrmMain As frmMain
Private Declare Function ActivateKeyboardLayout& Lib "user32" (ByVal HKL As Long, _
ByVal flags As Long)


Private Sub cmdSearch_Click()
Dim SS As String, S1 As String
Dim item As ListItem

ListView1.ListItems.Clear
SS = "SELECT * FROM dep WHERE "
If osob = True Then
    S1 = "dep_or Like " & Chr(34) & txtTarget & "%" & Chr(34) & " ORDER by dep_or;"
Else
    S1 = "dep_pib Like " & Chr(34) & txtTarget & "%" & Chr(34) & " ORDER by dep_pib;"
End If

ConnectToDataBase

myRS.Open SS & S1, myADO, adOpenDynamic

Do While Not myRS.EOF
    Set item = ListView1.ListItems.Add(, , myRS("dep_from"))
    item.SubItems(1) = (myRS("dep_or"))
    item.SubItems(2) = (myRS("dep_pib"))
    item.SubItems(3) = (myRS("dep_date"))
    item.SubItems(4) = (myRS("dep_ins"))
    myRS.MoveNext
Loop
End Sub

Private Sub Form_Load()
'osob = True
lblTarget = "Ведіть № ОР"
txtTarget.MaxLength = 6
'txtTarget.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set myRS = Nothing
Set myADO = Nothing
Unload Me
End Sub

Private Sub osob_Click()
ListView1.ListItems.Clear
lblTarget = "Ведіть № ОР"
With txtTarget
    .Text = ""
'    .SetFocus
End With
End Sub

Private Sub pib_Click()
ListView1.ListItems.Clear
lblTarget = "Ведіть ПІБ пенсіонера"

With txtTarget
    .Text = ""
    .SetFocus
End With
ActivateKeyboardLayout &H4220422, 3
End Sub

Private Sub txtTarget_KeyDown(KeyCode As Integer, Shift As Integer)

If osob = True Then
    txtTarget.Locked = IIf((KeyCode > 47 And KeyCode < 58) Or _
    (KeyCode > 95 And KeyCode < 107) Or _
    (KeyCode = 8) Or (KeyCode = 46) Or (KeyCode = 188), IIf(KeyCode = 188, _
    IIf(InStr(1, txtTarget, ",") = 0 And txtTarget.SelStart <> 0, False, True), False), True)
End If
If pib = True Then
    txtTarget.Locked = False
    txtTarget.MaxLength = 25
End If

End Sub
