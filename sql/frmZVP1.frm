VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmZVP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Звіт по вигрузках в розрізі дат "
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5088.607
   ScaleMode       =   0  'User
   ScaleWidth      =   6701.794
   Begin VB.ComboBox cmbYear1 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Text            =   "2010"
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox cmbMonth1 
      Height          =   315
      ItemData        =   "frmZVP1.frx":0000
      Left            =   480
      List            =   "frmZVP1.frx":0002
      TabIndex        =   0
      Text            =   "1"
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox cmbYear2 
      Height          =   315
      Left            =   4680
      TabIndex        =   3
      Text            =   "2010"
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox cmbMonth2 
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Text            =   "1"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   3360
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Формувати"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4260
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Дата"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Інспектор"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Кількість"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Загальна кількість"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Кінцева дата (місяць, рік)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Початкова дата (місяць, рік)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmZVP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public myFrmMain As frmMain

Private Sub cmbMonth1_Click()
If cmbMonth1 = 12 Then
    cmbMonth2 = 1
    cmbYear2 = cmbYear1 + 1
End If

End Sub

Private Sub cmdSelect_Click()
Dim zapros As String, _
      Item As ListItem, _
        AA, Beg, TheEnD
        

If Len(cmbMonth1) <> 2 Then cmbMonth1 = "0" & cmbMonth1
If Len(cmbMonth2) <> 2 Then cmbMonth2 = "0" & cmbMonth2

zapros = "SELECT Format([dep_date]," & Chr(34) & "mmmm yyyy" & Chr(34) & ") AS Dataa, dep_ins, Count(dep_ins) AS Kol " & _
"From Dep GROUP BY Format([dep_date]," & Chr(34) & "mmmm yyyy" & Chr(34) & "), dep_ins,  Format([dep_date]," & Chr(34) & "mm.yyyy" & Chr(34) & ") " & _
"HAVING (Format([dep_date]," & Chr(34) & "mm.yyyy" & Chr(34) & _
") Between '" & cmbMonth1 & "." & cmbYear1 & "' And '" & cmbMonth2 & "." & cmbYear2 & "')" & _
"ORDER BY Format([dep_date]," & Chr(34) & "mm.yyyy" & Chr(34) & "), dep_ins;"

ConnectToDataBase

myRS.Open zapros, myADO, adOpenStatic

ListView1.ListItems.Clear

Do While Not myRS.EOF
    Set Item = ListView1.ListItems.Add(, , myRS("Dataa"))
    Item.SubItems(1) = (myRS("dep_ins"))
    Item.SubItems(2) = (myRS("Kol"))
'    Text1.Text = Text1.Text + myRS("Kol")
    AA = AA + myRS("Kol")
    myRS.MoveNext
Loop
Text1.Text = AA

End Sub

Private Sub Form_Load()
Dim iMonth As Integer, iYear As Integer

For iMonth = 1 To 12
    cmbMonth1.AddItem iMonth
    cmbMonth2.AddItem iMonth
Next
    
For iYear = 2010 To 2020
    cmbYear1.AddItem iYear
    cmbYear2.AddItem iYear
Next

cmbMonth1 = Month(Date)
If cmbMonth1 = 12 Then
    cmbMonth2 = 1
    cmbYear2 = Year(Date) + 1
Else
    cmbMonth2 = cmbMonth1 + 1
End If

cmbYear1 = Year(Date)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
Set myRS = Nothing
Set myADO = Nothing
End Sub
