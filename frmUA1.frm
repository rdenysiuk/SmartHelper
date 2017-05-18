VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Реєструвати оновлення АСОПД"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   Icon            =   "frmUA1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check_ToRaj 
      Caption         =   "Тестування скинуто на райони"
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   4200
      Width           =   3855
   End
   Begin MSMask.MaskEdBox txt_SetUp 
      Height          =   345
      Left            =   1800
      TabIndex        =   15
      Top             =   2280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   16
      Mask            =   "##.##.#### ##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txt_Recieve 
      Height          =   345
      Left            =   1800
      TabIndex        =   13
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   16
      Mask            =   "##.##.#### ##:##"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Терміни тестування "
      DragIcon        =   "frmUA1.frx":000C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   1800
      TabIndex        =   9
      Top             =   2760
      Width           =   4335
      Begin MSMask.MaskEdBox txt_Draj 
         Height          =   345
         Left            =   1320
         TabIndex        =   6
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16744576
         MaxLength       =   16
         Mask            =   "##.##.#### ##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt_Dkiev 
         Height          =   345
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   8421631
         MaxLength       =   16
         Mask            =   "##.##.#### ##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Дата на Київ"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Дата районів"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmd_OpF 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   600
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "Додати"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CheckBox Check_Test 
      Caption         =   "Тестування"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txt_Inv 
      Height          =   975
      Left            =   1800
      MaxLength       =   150
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   4335
   End
   Begin VB.TextBox txt_NArh 
      Height          =   350
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Дата встановлення"
      Height          =   465
      Left            =   240
      TabIndex        =   14
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Дата прийому"
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Короткий опис"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Назва архіва розсилки"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   200
      Width           =   1455
   End
End
Attribute VB_Name = "frmUA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public myFrmMain As frmMain

Private Sub Check_Test_Click()

If Check_Test = 0 Then
    Prim_Test (False)
Else
    Prim_Test (True)
End If

End Sub

Private Sub Check_ToRaj_Click()
If Check_ToRaj = 0 Then
    Check_ToRaj.FontBold = False
Else
    Check_ToRaj.FontBold = True
End If
End Sub

Private Sub cmd_OpF_Click()
Dim sFile As String
With Dialog1
    If Len(Month(Date)) = 1 Then
        .InitDir = "W:\new_Asopdw\" & Year(Date) & "\0" & Month(Date)
    Else
        .InitDir = "W:\new_Asopdw\" & Year(Date) & "\" & Month(Date)
    End If
    
    .DialogTitle = "Оберіть архів розсилки"
    .Filter = "Архів оновлення (*.arj)|*.arj| Усі файли (*.*)|*.*"
    .ShowOpen
End With
With txt_NArh
    .Text = Dialog1.FileTitle
    .SetFocus
    .SelStart = 0
    .SelLength = Len(.Text)
End With

If Len(Dialog1.FileName) <> 0 Then
    sFile = Dialog1.FileName
    txt_recieve.Text = Left(FileSystem.FileDateTime(sFile), 16)
    txt_SetUp = Left(txt_recieve.Text, 10) & " " & Left(DateAdd("n", 24, Right(txt_recieve.Text, 5)), 5)
    cmdADD.Enabled = True
End If

End Sub

Private Sub cmdADD_Click()
Dim sql, sql_search, sUpd_Id, sql_add_test, Quest, Quest_T

If Check_Test.Value = 1 And Check_ToRaj.Value = 1 Then _
sql = "INSERT INTO UP_DATE (upd_Narh, upd_inv, upd_test, upd_prog, upd_recieve, upd_setup, upd_toraj) VALUES (" & _
      Chr(39) & txt_NArh.Text & "'," & _
      Chr(39) & txt_inv.Text & "'," & _
      1 & Chr(44) & _
      ReadINI("viezd", "ID", PathFileIni) & "," & _
      Chr(39) & txt_recieve.Text & "'," & _
      Chr(39) & txt_SetUp.Text & "'," & 1 & ")"
'End If
If Check_Test.Value = 1 And Check_ToRaj.Value = 0 Then _
    sql = "INSERT INTO UP_DATE (upd_Narh, upd_inv, upd_test, upd_prog, upd_recieve, upd_setup) VALUES (" & _
    Chr(39) & txt_NArh.Text & "'," & _
    Chr(39) & txt_inv.Text & "'," & _
    1 & Chr(44) & _
    ReadINI("viezd", "ID", PathFileIni) & Chr(44) & _
    Chr(39) & txt_recieve.Text & "'," & _
    Chr(39) & txt_SetUp.Text & "')"

If Check_Test.Value = 0 Then _
      sql = "INSERT INTO UP_DATE (upd_Narh, upd_inv, upd_prog, upd_recieve, upd_setup) VALUES (" & _
      Chr(39) & txt_NArh.Text & Chr(39) & Chr(44) & _
      Chr(39) & txt_inv.Text & Chr(39) & Chr(44) & _
      ReadINI("viezd", "ID", PathFileIni) & Chr(44) & _
      Chr(39) & txt_recieve.Text & Chr(39) & Chr(44) & _
      Chr(39) & txt_SetUp.Text & Chr(39) & ")"
    
ConnectToDataBase
myRS.Open sql, myADO, adOpenDynamic

If Check_Test.Value = 1 Then
    sql_search = "SELECT upd_id FROM up_date WHERE upd_narh = " & _
        Chr(39) & txt_NArh & Chr(39) & ";"
    myRS.Open sql_search, myADO, adOpenDynamic
    sUpd_Id = myRS("upd_id")
    
    myRS.close
    
    sql_add_test = "INSERT INTO test (tst_id, tst_dkiev, tst_draj) VALUES (" & _
        sUpd_Id & Chr(44) & _
        Chr(39) & txt_DKiev.Text & Chr(39) & Chr(44) & _
        Chr(39) & txt_DRaj.Text & Chr(39) & ");"
    
    myRS.Open sql_add_test, myADO, adOpenDynamic
    Quest_T = MsgBox("Тестову розсилку додано" & vbCrLf & vbCrLf & "Продовжити реєстрацію оновлень? - <Да>" & _
        vbCrLf & "Завершити реєстрацію? - <Нет>", vbExclamation + vbYesNo, "Attention")
    If Quest_T = vbYes Then
    txt_NArh = ""
    txt_inv = ""
    txt_DKiev.Text = "__.__.____ __:__"
    txt_DRaj.Text = "__.__.____ __:__"
    Call cmd_OpF_Click
    Check_Test.Value = 0
    Else
    Unload Me
    End If
Else
    Quest = MsgBox("Розсилку оновлення додано" & vbCrLf & vbCrLf & "Продовжити реєстрацію оновлень? - <Да>" & _
    vbCrLf & "Завершити реєстрацію? - <Нет>", vbExclamation + vbYesNo, "Attention")
    If Quest = vbYes Then
    txt_NArh = ""
    txt_inv = ""
    Call cmd_OpF_Click
    Else
    Unload Me
    End If
End If

End Sub

Private Sub Form_Load()
With Me
.Width = 6500
.Height = 5300
End With

If Check_Test = 0 Then
    Prim_Test (False)
Else
    Prim_Test (True)
End If
Label1.Caption = "Назва архіва" & vbCrLf & "розсилки"
Check_ToRaj.Enabled = False
cmdADD.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
myFrmMain.reg_upd.Enabled = True
Unload Me
'Set myRS = Nothing
End Sub


Private Sub txt_NArh_Validate(Cancel As Boolean)
Dim Sql_Search_Narh
ConnectToDataBase
Sql_Search_Narh = "SELECT * FROM up_date WHERE upd_narh = " & Chr(39) & txt_NArh.Text & Chr(39)
myRS.Open Sql_Search_Narh, myADO, adOpenStatic
If myRS.RecordCount > 0 Then
    MsgBox txt_NArh.Text & " розсилка зареєстрована!!" & vbCrLf, vbCritical
    cmdADD.Enabled = False
Else
    cmdADD.Enabled = True
End If
End Sub


Public Function Prim_Test(Status As Boolean)
If Status = True Then
  Frame1.Enabled = True
  Label3.Enabled = True
  Label4.Enabled = True
  Check_Test.FontBold = True
  Check_ToRaj.Enabled = True
  txt_DKiev.Text = Date & " " & Left(Time, 5)
  txt_DRaj.Text = Date & " " & Left(Time, 5)
  'MsgBox Date & Time
Else
  Frame1.Enabled = False
  Label3.Enabled = False
  Label4.Enabled = False
  Check_Test.FontBold = False
  txt_DKiev.Text = "__.__.____ __:__"
  txt_DRaj.Text = "__.__.____ __:__"
  
  With Check_ToRaj
  .Value = 0
  .Enabled = False
  End With
End If

End Function
