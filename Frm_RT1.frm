VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRT 
   Caption         =   "Результати районів по тестуванню"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11385
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_RT1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9705
   ScaleWidth      =   11385
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_recieve 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4680
      TabIndex        =   27
      Top             =   280
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridResult 
      Height          =   6135
      Left            =   360
      TabIndex        =   22
      Top             =   2160
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   10821
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   -2147483630
      Rows            =   3
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   16777215
      BackColorSel    =   8388608
      ForeColorSel    =   16777215
      GridColor       =   8388608
      GridColorFixed  =   8388608
      AllowBigSelection=   0   'False
      HighLight       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).BandIndent=   1
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0).ColHeader=   1
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      Caption         =   " Короткий опис "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   6600
      TabIndex        =   19
      Top             =   240
      Width           =   4215
      Begin VB.TextBox txt_inv 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      Caption         =   " Зареєстровані результати "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   6735
      Left            =   240
      TabIndex        =   18
      Top             =   1800
      Width           =   6015
      Begin VB.CommandButton cmd_ZV1 
         Caption         =   "Звіт по результатах"
         Height          =   615
         Left            =   4200
         TabIndex        =   25
         Top             =   4320
         Width           =   1695
      End
      Begin VB.CommandButton cmd_ZV2 
         Caption         =   "Розширений звіт"
         Height          =   615
         Left            =   4200
         TabIndex        =   24
         ToolTipText     =   "Звіт по тестуванню кількох результатах на одне тестування"
         Top             =   5040
         Width           =   1695
      End
      Begin VB.CommandButton cmd_EndTest 
         Caption         =   "Закрити тестування"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   4200
         TabIndex        =   23
         ToolTipText     =   "Закривати тестування розсилки суворо лише при занесених всіх районних результатах"
         Top             =   5880
         Width           =   1665
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Додавання результатів"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   6735
      Left            =   6600
      TabIndex        =   14
      Top             =   1800
      Width           =   4215
      Begin MSMask.MaskEdBox txt_Dzvit 
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         ToolTipText     =   "Заповнюється у випадку кількох результатів на одне тестування"
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.##.####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmbRaj 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   800
         Width           =   1935
      End
      Begin VB.ComboBox cmbComment 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   1250
         Width           =   1935
      End
      Begin VB.CheckBox chckLate 
         Caption         =   "Запізнення"
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
         Left            =   1680
         TabIndex        =   5
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton cmbAdd 
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
         Left            =   1080
         TabIndex        =   7
         Top             =   6000
         Width           =   2655
      End
      Begin VB.TextBox txtPrim 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   120
         MaxLength       =   400
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   2280
         Width           =   3855
      End
      Begin VB.Label Label7 
         Caption         =   "Звітна дата"
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
         Left            =   600
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Короткий зміст листа"
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
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Район"
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
         Left            =   1000
         TabIndex        =   16
         Top             =   795
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Тип зауваження"
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
         Left            =   120
         TabIndex        =   15
         Top             =   1245
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   " Терміни тестування "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   6015
      Begin VB.TextBox txt_DRaj 
         BackColor       =   &H00FF8080&
         Height          =   300
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txt_DKiev 
         BackColor       =   &H008080FF&
         Height          =   300
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Дата на райони"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Дата на Київ"
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.ComboBox cmbTest 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "від"
      Height          =   300
      Left            =   4320
      TabIndex        =   26
      Top             =   280
      Width           =   225
   End
   Begin VB.Label lblKodTest 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Тестова розсилка"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "FrmRT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public myFrmMain As frmMain
Public KolResTest As Integer

Private Sub cmbAdd_Click()
Dim InsertIn, myLate
Dim myDataZvit As String

myDataZvit = "NULL"
cmbTest.Enabled = False
myLate = 0

If chckLate.Value = 1 Then myLate = 1

If txt_Dzvit.Text <> "__.__.____" Then
    myDataZvit = "'" & txt_Dzvit.Text & "'"
Else
    myDataZvit = "'" & Date & "'"
End If

ConnectToDataBase
InsertIn = "INSERT INTO result (res_dzvit, res_upd, res_raj, res_comment, res_late, res_prim) " & _
    "VALUES (" & myDataZvit & "," & _
    lblKodTest.Caption & "," & _
    cmbRaj & "," & _
    cmbComment.ItemData(cmbComment.ListIndex) & "," & _
    myLate & ",'" & txtPrim.Text & "');"
myRS.Open InsertIn, myADO, adOpenStatic
MsgBox "Результат по району - " & cmbRaj & " додано", , "Attention"
txtPrim.Text = ""
If chckLate.Value = 1 Then chckLate.Value = 0

With cmbRaj
    If .Text <> 6826 Then .ListIndex = .ListIndex + 1
    '.Text = ""
    .SetFocus
End With
Call cmbTest_Validate(True)

End Sub


Private Sub cmbTEST_Click()
Dim SelList
With cmbTest
lblKodTest.Caption = .ItemData(.ListIndex)
End With
ConnectToDataBase
SelList = "SELECT upd_inv, upd_recieve, tst_dkiev, tst_draj FROM up_date " & _
    "INNER JOIN test ON upd_id = tst_id WHERE upd_test = 1 and upd_narh = '" & cmbTest & "';"
myRS.Open SelList, myADO, adOpenStatic
txt_inv.Text = myRS("upd_inv")
txt_DKiev = Left(myRS("tst_dkiev"), 16)
txt_DRaj = Left(myRS("tst_draj"), 16)
txt_recieve = Left(myRS("upd_recieve"), 10)


End Sub

Private Sub cmbTest_Validate(Cancel As Boolean)
Dim SelResultTest
Dim I As Integer
Dim iIndx As Integer

ConnectToDataBase
SelResultTest = "SELECT convert(varchar(10),res_dzvit,104) as res_dzvit, res_raj, com_name, (case when res_late = 0 then '' else 'ТАК' end) as late FROM result " & _
    "INNER JOIN comment ON res_comment=com_id WHERE res_upd = " & lblKodTest & " ORDER BY res_dzvit, res_raj"

If Len(lblKodTest.Caption) > 0 Then
  myRS.Open SelResultTest, myADO, adOpenStatic
  
  If myRS.RecordCount <> 0 Then
    Set GridResult.DataSource = myRS
    With GridResult
    .Row = 0
    .Col = iIndx
    .CellAlignment = 4
    .MergeCol(iIndx) = True
    .Col = 0
    .ColSel = .Cols - 1
    .MergeCells = flexMergeRestrictColumns
    End With
    cmd_EndTest.Enabled = True
    KolResTest = myRS.RecordCount
  End If
  
  If myRS.RecordCount = 0 Then
   Call SetGridData
   cmd_EndTest.Enabled = False
   cmd_ZV1.Enabled = False
   cmd_ZV2.Enabled = False
  Else
   cmd_ZV1.Enabled = True
   cmd_ZV2.Enabled = True
  End If
  
End If

End Sub


Private Sub cmd_EndTest_Click()
Dim QEndTest, SetEndTest As String
QEndTest = MsgBox("Якщо закрити тестування розсилки, в подальшому не можливо " & vbCrLf & "буде сформувати звіт з результатами." & vbCrLf & _
    "Перш ніж закрити тестування, впевніться в наявності звіту." & _
    vbCrLf & vbCrLf & "Дійсно бажаєте закрити тестування по розсилці " & cmbTest & ": " _
    & vbCrLf & Left(txt_inv.Text, 55) & " ...", vbExclamation + vbYesNo + vbDefaultButton2, "Warning")
If QEndTest = vbYes Then
   SetEndTest = "UPDATE up_date SET upd_end=1 WHERE upd_id=" & lblKodTest.Caption
   ConnectToDataBase
   myRS.Open SetEndTest, myADO, adOpenDynamic
   
   'Call Form_Load
   'cmbTest.ListIndex = cmbTest.ListIndex + 1
   'Call cmbTest_Validate(True)
   Call Form_Load
   txt_recieve.Text = ""
   txt_DKiev.Text = ""
   txt_DRaj.Text = ""
   txt_inv.Text = ""
   With GridResult
   .Clear
   .Rows = 3
   .Cols = 4
   End With
   Call SetGridData
   cmbTest.SetFocus
   
End If
End Sub

Private Sub cmd_ZV1_Click()
Dim SelResult As String
Dim xlAppl As Excel.Application
Dim xlBook  As Excel.Workbook

Set xlAppl = Excel.Application
Set xlBook = xlAppl.Workbooks.Open(App.Path & "\Report\RepResTest1.xlt")

xlAppl.Visible = True

SelResult = "SELECT res_raj, oper_nraj, " & _
"sum((case when res_late = 0 then com_perc else com_perc - 30 end))/count((case when res_late = 0 then com_perc else com_perc - 30 end)) as mproc " & _
"FROM result INNER JOIN comment ON res_comment=com_id INNER JOIN oper ON res_raj=oper_raj " & _
"Where res_upd = " & lblKodTest.Caption & " GROUP BY res_raj, oper_nraj ORDER BY res_raj"

ConnectToDataBase
myRS.Open SelResult, myADO, adOpenStatic

With xlBook.Worksheets(1)
.Range("b11").CopyFromRecordset myRS
.Range("c3").Value = cmbTest
.Range("c6").Value = txt_inv.Text
.Range("c4").Value = txt_recieve.Text
.Range("c8").Value = txt_DRaj.Text
.Range("d36").Value = myFrmMain.StatusBar1.Panels(8).Text
End With
'MsgBox myRS.RecordCount
End Sub

Private Sub cmd_ZV2_Click()
Dim SelResAdvanced As String
Dim SelResAdvanced1 As String
Dim SelResAdvanced2 As String

Dim xlAppl As Excel.Application
Dim xlBook  As Excel.Workbook

Set xlAppl = Excel.Application

xlAppl.Visible = True

ConnectToDataBase

'If GridResult.Rows = 25 Then
    SelResAdvanced = "SELECT res_raj, oper_nraj, com_name, " & _
                    "(case when res_late = 0 then '' else 'ТАК' end) as late " & _
                    "FROM sspz.dbo.result  INNER JOIN sspz.dbo.comment ON res_comment=com_id INNER JOIN sspz.dbo.oper ON res_raj=oper_raj " & _
                    "WHERE res_upd =" & lblKodTest.Caption & " ORDER BY res_raj;"
    
    myRS.Open SelResAdvanced, myADO, adOpenDynamic
    Set xlBook = xlAppl.Workbooks.Open(App.Path & "\Report\RepResTest2.xlt")
    With xlBook.Worksheets(1)
    .Range("b12").CopyFromRecordset myRS
    .Range("c3").Value = cmbTest
    .Range("c6").Value = txt_inv.Text
    .Range("c4").Value = txt_recieve.Text
    .Range("c8").Value = txt_DRaj.Text
    .Range("d37").Value = myFrmMain.StatusBar1.Panels(8).Text
    End With
     'MsgBox SelResAdvanced
'Else
'    SelResAdvanced = _
'    "SELECT res_dzvit, [6801], [6802], [6803], [6804], [6805], [6806], [6807], [6808], [6809], " & _
'    "[6810], [6811], [6812], [6813], [6814], [6815], [6816], [6817], [6818], [6819], [6820], " & _
'    " [6821], [6823], [6824], [6826] FROM "
'    SelResAdvanced1 = _
'    "(SELECT oper_raj, res_dzvit, (case when res_late = 1 then com_perc-30 else com_perc end) as pr " & _
'    "FROM result AS R JOIN comment AS C ON R.res_comment=C.com_id JOIN oper AS OO ON R.res_raj=OO.oper_raj " & _
'    "WHERE res_upd = " & lblKodTest.Caption & " AS ZZ " & _
'    "PIVOT (max([PR]) for [oper_raj]"
'    SelResAdvanced2 = _
'    " in ([6801], [6802], [6803], [6804], [6805], [6806], [6807], [6808], [6809], [6810], " & _
'    "[6811], [6812], [6813], [6814], [6815], [6816], [6817], [6818], [6819], [6820], [6821], [6823], [6824], " & _
'    "[6826])) as pvt"
'
'    myRS.Open SelResAdvanced & SelResAdvanced1 & SelResAdvanced2, myADO, adOpenDynamic
'    Set xlBook = xlAppl.Workbooks.Open(App.Path & "\Report\RepResTest3.xlt")
'    With xlBook.Worksheets(1)
'    .Range("b11").CopyFromRecordset myRS
'    .Range("c3").Value = cmbTest
'    .Range("c6").Value = txt_inv.Text
'    .Range("c4").Value = txt_recieve.Text
'    .Range("c8").Value = txt_DRaj.Text
'    .Range("l25").Value = myFrmMain.StatusBar1.Panels(8).Text
'    End With
'End If



End Sub



Private Sub Form_Load()
Dim SelTest, SelRaj, SelComment

cmd_EndTest.Enabled = False
cmd_ZV1.Enabled = False
cmd_ZV2.Enabled = False

ConnectToDataBase

SelTest = "SELECT upd_narh, upd_id FROM up_date WHERE upd_test=1 and upd_toraj=1 and (upd_end is null or upd_end =0)"
myRS.Open SelTest, myADO, adOpenStatic
With cmbTest
.Clear
Do While Not myRS.EOF
    .AddItem myRS("upd_narh").Value
    .ItemData(.NewIndex) = myRS("upd_id")
    myRS.MoveNext
Loop
End With

myRS.close

SelRaj = "SELECT oper_raj FROM oper WHERE oper_raj <> 6825 ORDER BY oper_raj"
myRS.Open SelRaj, myADO, adOpenStatic

Do While Not myRS.EOF
    cmbRaj.AddItem myRS("oper_raj").Value
    myRS.MoveNext
Loop

myRS.close
SelComment = "SELECT com_id, com_name FROM comment"
myRS.Open SelComment, myADO, adOpenStatic

Do While Not myRS.EOF
    With cmbComment
    .AddItem myRS("com_name").Value
    .ItemData(.NewIndex) = myRS("com_id").Value
    End With
    myRS.MoveNext
Loop
txt_Dzvit.Text = Date

Call SetGridData

End Sub

Private Sub SetGridData()
With GridResult
    .Rows = 3
    .Cols = 4
    .ColWidth(0) = 1000
    .ColHeaderCaption(0, 0) = "Звітна дата"
    .ColWidth(1) = 600
    .ColHeaderCaption(0, 1) = "Район"
    .ColWidth(2) = 1200
    .ColHeaderCaption(0, 2) = "Зауваження"
    .ColWidth(3) = 850
    .ColHeaderCaption(0, 3) = "Невчасно"
End With
End Sub

