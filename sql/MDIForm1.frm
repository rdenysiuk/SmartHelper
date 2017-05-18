VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H80000003&
   Caption         =   "MDIform1"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11880
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8940
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1667
            MinWidth        =   1658
            Text            =   "Дата:"
            TextSave        =   "Дата:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "09.11.2010"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   3413
            MinWidth        =   3422
            Text            =   "Шлях до бази даних"
            TextSave        =   "Шлях до бази даних"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12726
         EndProperty
      EndProperty
   End
   Begin VB.Menu work 
      Caption         =   "Робота"
      Begin VB.Menu viezd 
         Caption         =   "Виїзд пенсіонера"
      End
      Begin VB.Menu RepSend 
         Caption         =   "Звіти на район"
      End
      Begin VB.Menu null 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Вихід"
      End
   End
   Begin VB.Menu report 
      Caption         =   "Звіти"
      Begin VB.Menu ArhRep 
         Caption         =   "Вигрузки 1ls"
      End
      Begin VB.Menu find1ls 
         Caption         =   "Пошук вигрузки 1ls"
      End
   End
   Begin VB.Menu Setting 
      Caption         =   "Налаштування"
   End
   Begin VB.Menu window 
      Caption         =   "Вікна"
      WindowList      =   -1  'True
   End
   Begin VB.Menu About 
      Caption         =   "Довідка"
      Begin VB.Menu AboutProg 
         Caption         =   "Про програму"
      End
   End
   Begin VB.Menu close 
      Caption         =   "Вихід"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myFrmSetting As New frmSetting
Private Sub AboutProg_Click()
MsgBox "Програмний комплекс " & Chr(34) & "SmartHelper" & Chr(34) & "." & _
        vbCrLf & "Створений для автоматизації робочого місця спеціаліста відділу ССПЗ" & _
        vbCrLf & "ГУ ПФУ в Хмельницькій області." & _
        vbCrLf & vbCrLf & "Автор комплексу: Денисюк Роман" & _
        vbCrLf & "E-mail: denisyik_roma@ukr.net" & _
        vbCrLf & "ICQ: 426801527", vbInformation, "Про програму..."
End Sub

Private Sub ArhRep_Click()
Dim myFrmZVp As New frmZVP
Set myFrmZVp.myFrmMain = Me
myFrmZVp.Show
End Sub

Private Sub close_Click()
Unload Me
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub find1ls_Click()
Dim myFrmFindVP As New frmFindVP
Set myFrmFindVP.myFrmMain = Me
myFrmFindVP.Show
End Sub

Private Sub MDIForm_Load()
Dim FO As New FileSystemObject
frmMain.Caption = "SmartHelper v2.0    " & Chr(169) & " Romario"
DataBasePath = ReadINI("database", "dbpath", PathFileIni)

If Not FO.FileExists(DataBasePath) Then
    MsgBox DataBasePath, vbCritical, "Відсутня база данних"
    myFrmSetting.Show
    myFrmSetting.cmbSelFolderBD.SetFocus
Else
    StatusBar1.Panels(4).Text = DataBasePath
End If

End Sub

Private Sub RepSend_Click()
Dim myFrmSR As New frmSR
Set myFrmSR.myFrmMain = Me
myFrmSR.Show
'work.Enabled = False
RepSend.Enabled = False
'report.Enabled = False
End Sub

Private Sub Setting_Click()
Set myFrmSetting.myFrmMain = Me
myFrmSetting.Show
Setting.Enabled = False

End Sub

Private Sub viezd_Click()
Dim myFrmVp As New frmVP
Set myFrmVp.myFrmMain = Me
myFrmVp.Show
'work.Enabled = False
viezd.Enabled = False
'report.Enabled = False
End Sub

