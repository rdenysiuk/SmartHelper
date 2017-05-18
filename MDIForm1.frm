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
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   8925
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1773
            MinWidth        =   1763
            Text            =   "Дата:"
            TextSave        =   "Дата:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "16.11.2011"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   3175
            MinWidth        =   3174
            Text            =   "Сервер"
            TextSave        =   "Сервер"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   3969
            MinWidth        =   3968
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   2196
            MinWidth        =   2187
            Text            =   "З'єднання"
            TextSave        =   "З'єднання"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   2187
            MinWidth        =   2187
            Text            =   "Користувач:"
            TextSave        =   "Користувач:"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2884
            MinWidth        =   2893
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
      Begin VB.Menu ndi 
         Caption         =   "НДІ"
      End
      Begin VB.Menu null1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Вихід"
      End
   End
   Begin VB.Menu teh 
      Caption         =   "Обслуговування техніки"
      Begin VB.Menu addWork 
         Caption         =   "Додати роботу"
      End
      Begin VB.Menu dov 
         Caption         =   "Довідник найменувань"
      End
   End
   Begin VB.Menu update1 
      Caption         =   "Оновлення АСОПД"
      Begin VB.Menu reg_upd 
         Caption         =   "Реєстрація розсилок"
      End
      Begin VB.Menu result_test 
         Caption         =   "Результати тестувань"
      End
   End
   Begin VB.Menu report 
      Caption         =   "Звітність"
      Begin VB.Menu ArhRep 
         Caption         =   "Виїзди пенсіонерів"
      End
      Begin VB.Menu upd_set 
         Caption         =   "Журнал оновлень АСОПД"
      End
      Begin VB.Menu zndi 
         Caption         =   "Нормативно-довідкова інфо"
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
With App
MsgBox "Програмний комплекс " & Chr(34) & "SmartHelper v" & .Major & "." & .Minor & "." & .Revision & _
    Chr(34) & "." & vbCrLf & "I am too lazy to do this work in manual mode" & vbCrLf & _
    vbCrLf & "Створений для автоматизації робочого місця спеціаліста відділу ССПЗ" & _
    vbCrLf & "ГУ ПФУ в Хмельницькій області." & _
    vbCrLf & vbCrLf & "Автор: Денисюк Роман" & _
    vbCrLf & "E-mail: denisyik_roma@ukr.net" & _
    vbCrLf & "ICQ: 426801527", vbInformation, "Про програму..."
End With
End Sub

Private Sub addWork_Click()
Dim myFrmAddWork As New frmAddWork
Set myFrmAddWork.myFrmMain = Me
myFrmAddWork.Show
addWork.Enabled = False
End Sub

Private Sub ArhRep_Click()
Dim myFrmZVp As New frmZVP
Set myFrmZVp.myFrmMain = Me
myFrmZVp.Show
End Sub

Private Sub close_Click()
Unload Me
End Sub

Private Sub dov_Click()
Dim myFrmAddType As New frmAddType
Set myFrmAddType.myFrmMain = Me
myFrmAddType.Show
dov.Enabled = False
End Sub

Private Sub exit_Click()
Unload Me
End Sub


Private Sub MDIForm_Load()
Dim Sel_pib
With App
frmMain.Caption = "SmartHelper v" & .Major & "." & .Minor & "   " & Chr(169) & " Romario"
End With

ConnectToDataBase
Sel_pib = "SELECT pr_pib FROM progs WHERE pr_id = " & ReadINI("viezd", "ID", PathFileIni)
myRS.Open Sel_pib, myADO, adOpenStatic

With StatusBar1
    .Panels(4).Text = Left(myADO.Properties("Data Source").Value, InStr(myADO.Properties("Data Source").Value, "\") - 1)
    If myRS.RecordCount >= 1 Then
        .Panels(8).Text = myRS("pr_pib")
    Else
        work.Enabled = False
        report.Enabled = False
        teh.Enabled = False
        update1.Enabled = False
        MsgBox "У файлі конфігурації не коректні/відсутні дані користувача." & vbCrLf & vbCrLf & _
            "Введіть свій ідентифікатор, прізвище та ініціали." & vbCrLf & _
            "Після цього перезапустіть програму", vbCritical + vbOKOnly, "Error"
        Call Setting_Click
        
        With myFrmSetting
            .txtVP(4).SetFocus
            .txtVP(4).SelStart = 0
            .txtVP(4).SelLength = Len(.txtVP(4).Text)
        End With
        
    End If
    
    If myADO.State = 1 Then
        .Panels(6).Text = "Active"
    Else
        .Panels(6).Text = "NO CONNECT!!! ERROR"
    End If
    
End With

'teh.Visible = False
End Sub

Private Sub ndi_Click()
Dim myFrmNDI As New frmNDI
Set myFrmNDI.myFrmMain = Me
myFrmNDI.Show
End Sub

Private Sub reg_upd_Click()
Dim myFrmUA As New frmUA
Set myFrmUA.myFrmMain = Me
myFrmUA.Show
reg_upd.Enabled = False
End Sub

Private Sub RepSend_Click()
Dim myFrmSR As New frmSR
Set myFrmSR.myFrmMain = Me
myFrmSR.Show
RepSend.Enabled = False
End Sub

Private Sub result_test_Click()
Dim myFrmRT As New FrmRT
Set myFrmRT.myFrmMain = Me
myFrmRT.Show
End Sub

Private Sub Setting_Click()
Set myFrmSetting.myFrmMain = Me
myFrmSetting.Show
Setting.Enabled = False

End Sub


Private Sub upd_set_Click()
Dim myFrmUpset As New frmUpSet
Set myFrmUpset.myFrmMain = Me
myFrmUpset.Show
End Sub

Private Sub viezd_Click()
Dim myFrmVp As New frmVP
Set myFrmVp.myFrmMain = Me
myFrmVp.Show
viezd.Enabled = False
End Sub

Private Sub zndi_Click()
Dim myFrmZndi As New frmZNDI
Set myFrmZndi.myFrmMain = Me
myFrmZndi.Show
zndi.Enabled = False
End Sub
