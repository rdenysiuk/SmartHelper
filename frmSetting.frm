VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetting 
   Caption         =   "Налаштування програми"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   10875
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   13573
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   556
      TabCaption(0)   =   "Загальне"
      TabPicture(0)   =   "frmSetting.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtBD"
      Tab(0).Control(1)=   "cmbSaveBD"
      Tab(0).Control(2)=   "cmbSelFolderBD"
      Tab(0).Control(3)=   "Label8"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Виїзд пенсіонера"
      TabPicture(1)   =   "frmSetting.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmbSaveVP"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtVP(4)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtVP(3)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmbSelFolder_VP"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtVP(0)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtVP(2)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "FTP з'єднання"
      TabPicture(2)   =   "frmSetting.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label15"
      Tab(2).Control(1)=   "Label13(1)"
      Tab(2).Control(2)=   "Label14"
      Tab(2).Control(3)=   "Label13(0)"
      Tab(2).Control(4)=   "Label12"
      Tab(2).Control(5)=   "Label10"
      Tab(2).Control(6)=   "Label11"
      Tab(2).Control(7)=   "Label9"
      Tab(2).Control(8)=   "txtFtp(7)"
      Tab(2).Control(9)=   "txtFtp(6)"
      Tab(2).Control(10)=   "txtFtp(5)"
      Tab(2).Control(11)=   "cmbSaveFtp"
      Tab(2).Control(12)=   "txtFtp(4)"
      Tab(2).Control(13)=   "txtFtp(3)"
      Tab(2).Control(14)=   "txtFtp(2)"
      Tab(2).Control(15)=   "txtFtp(1)"
      Tab(2).Control(16)=   "txtFtp(0)"
      Tab(2).Control(17)=   "cmbTEST(1)"
      Tab(2).Control(18)=   "cmbTEST(0)"
      Tab(2).ControlCount=   19
      TabCaption(3)   =   "Надіслати звіти"
      TabPicture(3)   =   "frmSetting.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label5"
      Tab(3).Control(1)=   "Label6"
      Tab(3).Control(2)=   "Label16"
      Tab(3).Control(3)=   "txtSR(0)"
      Tab(3).Control(4)=   "cmbSelFolder_SR1"
      Tab(3).Control(5)=   "txtSR(1)"
      Tab(3).Control(6)=   "cmbSelFolder_SR2"
      Tab(3).Control(7)=   "cmbSaveSR"
      Tab(3).Control(8)=   "Chk_Oper4"
      Tab(3).Control(9)=   "Chk_Oper3"
      Tab(3).Control(10)=   "Chk_Oper2"
      Tab(3).Control(11)=   "Chk_Oper1"
      Tab(3).ControlCount=   12
      Begin VB.TextBox txtVP 
         Height          =   315
         Index           =   2
         Left            =   6600
         TabIndex        =   45
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CheckBox Chk_Oper1 
         Caption         =   "1-й операційний"
         Height          =   495
         Left            =   -72600
         TabIndex        =   24
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CheckBox Chk_Oper2 
         Caption         =   "2-й операційний"
         Height          =   495
         Left            =   -72600
         TabIndex        =   25
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox Chk_Oper3 
         Caption         =   "3-й операційний"
         Height          =   495
         Left            =   -72600
         TabIndex        =   43
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CheckBox Chk_Oper4 
         Caption         =   "4-й операційний"
         Height          =   495
         Left            =   -72600
         TabIndex        =   26
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox txtBD 
         Height          =   315
         Left            =   -72700
         TabIndex        =   1
         Top             =   1000
         Width           =   5415
      End
      Begin VB.CommandButton cmbSaveBD 
         Caption         =   "Зберегти"
         Height          =   375
         Left            =   -69360
         TabIndex        =   3
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CommandButton cmbSelFolderBD 
         Caption         =   "..."
         Height          =   315
         Left            =   -67320
         TabIndex        =   2
         Top             =   1000
         Width           =   400
      End
      Begin VB.TextBox txtVP 
         Height          =   315
         Index           =   0
         Left            =   2280
         TabIndex        =   4
         Top             =   1000
         Width           =   5415
      End
      Begin VB.CommandButton cmbSelFolder_VP 
         Caption         =   "..."
         Height          =   315
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1000
         Width           =   400
      End
      Begin VB.TextBox txtVP 
         Height          =   315
         Index           =   3
         Left            =   2280
         TabIndex        =   6
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtVP 
         Height          =   315
         Index           =   4
         Left            =   2280
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmbTEST 
         Caption         =   "Тест з'єднання"
         Height          =   350
         Index           =   0
         Left            =   -68880
         TabIndex        =   13
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton cmbTEST 
         Caption         =   "Тест з'єднання"
         Height          =   350
         Index           =   1
         Left            =   -68880
         TabIndex        =   18
         Top             =   3650
         Width           =   1935
      End
      Begin VB.CommandButton cmbSaveVP 
         Caption         =   "Зберегти"
         Height          =   400
         Left            =   5640
         TabIndex        =   8
         Top             =   4080
         Width           =   2500
      End
      Begin VB.TextBox txtFtp 
         Height          =   315
         Index           =   0
         Left            =   -72700
         TabIndex        =   9
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtFtp 
         Height          =   315
         Index           =   1
         Left            =   -69720
         TabIndex        =   10
         Top             =   1000
         Width           =   2775
      End
      Begin VB.TextBox txtFtp 
         Height          =   315
         Index           =   2
         Left            =   -72700
         TabIndex        =   11
         Top             =   1400
         Width           =   2055
      End
      Begin VB.TextBox txtFtp 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   -69720
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   1400
         Width           =   2775
      End
      Begin VB.TextBox txtFtp 
         Height          =   315
         Index           =   4
         Left            =   -72700
         TabIndex        =   14
         Top             =   2880
         Width           =   2055
      End
      Begin VB.CommandButton cmbSaveFtp 
         Caption         =   "Зберегти"
         Height          =   400
         Left            =   -69120
         TabIndex        =   19
         Top             =   5520
         Width           =   2500
      End
      Begin VB.TextBox txtFtp 
         Height          =   315
         Index           =   5
         Left            =   -69720
         TabIndex        =   15
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox txtFtp 
         Height          =   315
         Index           =   6
         Left            =   -72720
         TabIndex        =   16
         Top             =   3240
         Width           =   2055
      End
      Begin VB.TextBox txtFtp 
         Height          =   315
         Index           =   7
         Left            =   -69720
         TabIndex        =   17
         Top             =   3240
         Width           =   2775
      End
      Begin VB.CommandButton cmbSaveSR 
         Caption         =   "Зберегти"
         Height          =   375
         Left            =   -69720
         TabIndex        =   27
         Top             =   4080
         Width           =   2535
      End
      Begin VB.CommandButton cmbSelFolder_SR2 
         Caption         =   "..."
         Height          =   315
         Left            =   -67680
         TabIndex        =   23
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtSR 
         Height          =   315
         Index           =   1
         Left            =   -72600
         TabIndex        =   22
         Top             =   1320
         Width           =   4935
      End
      Begin VB.CommandButton cmbSelFolder_SR1 
         Caption         =   "..."
         Height          =   315
         Left            =   -67680
         TabIndex        =   21
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtSR 
         Height          =   315
         Index           =   0
         Left            =   -72600
         TabIndex        =   20
         Top             =   960
         Width           =   4935
      End
      Begin VB.Label Label3 
         Caption         =   "№ телефону"
         Height          =   255
         Left            =   5520
         TabIndex        =   44
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Список районів"
         Height          =   255
         Left            =   -74280
         TabIndex        =   42
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Шлях до БД"
         Height          =   195
         Left            =   -73800
         TabIndex        =   41
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Каталог передачі"
         Height          =   195
         Left            =   720
         TabIndex        =   40
         Top             =   1000
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Прізвище та ініціали"
         Height          =   195
         Left            =   720
         TabIndex        =   39
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ідентифікатор"
         Height          =   195
         Left            =   1080
         TabIndex        =   38
         Top             =   1920
         Width           =   1125
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Ftp сервер (області)"
         Height          =   195
         Left            =   -74400
         TabIndex        =   37
         Top             =   1050
         Width           =   1530
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Ftp київський"
         Height          =   195
         Left            =   -73800
         TabIndex        =   36
         Top             =   3000
         Width           =   1020
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Каталог"
         Height          =   195
         Left            =   -70440
         TabIndex        =   35
         Top             =   2880
         Width           =   630
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Каталог"
         Height          =   195
         Left            =   -70440
         TabIndex        =   34
         Top             =   1050
         Width           =   630
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Логін"
         Height          =   195
         Index           =   0
         Left            =   -73200
         TabIndex        =   33
         Top             =   1500
         Width           =   390
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Пароль"
         Height          =   195
         Left            =   -70440
         TabIndex        =   32
         Top             =   1455
         Width           =   555
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Логін"
         Height          =   195
         Index           =   1
         Left            =   -73200
         TabIndex        =   31
         Top             =   3360
         Width           =   390
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Пароль"
         Height          =   195
         Left            =   -70440
         TabIndex        =   30
         Top             =   3240
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Поштовий каталог"
         Height          =   195
         Left            =   -74160
         TabIndex        =   29
         Top             =   1320
         Width           =   1425
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Каталог районних БД"
         Height          =   195
         Left            =   -74400
         TabIndex        =   28
         Top             =   960
         Width           =   1650
      End
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public myFrmMain As frmMain

'FTP
Dim hConnection As Long, _
          hOpen As Long, _
          hFile As Long, _
         dwType As Long, _
        dwSeman As Long


Private Sub Check1_Click()
cmbSaveSR.Enabled = True
End Sub

Private Sub Chk_Oper1_Click()
cmbSaveSR.Enabled = True
End Sub

Private Sub Chk_Oper2_Click()
cmbSaveSR.Enabled = True
End Sub

Private Sub Chk_Oper3_Click()
cmbSaveSR.Enabled = True
End Sub

Private Sub Chk_Oper4_Click()
cmbSaveSR.Enabled = True
End Sub

Private Sub cmbSaveBD_Click()
WriteINI "DataBase", "dbpath", txtBD.Text, PathFileIni
cmbSaveBD.Enabled = False
End Sub

Private Sub cmbSaveFtp_Click()
'ФТП обласний
WriteINI "ftpconnect", "ftpobl", txtFtp(0).Text, PathFileIni
WriteINI "ftpconnect", "oblkat", txtFtp(1).Text, PathFileIni
WriteINI "ftpconnect", "logobl", txtFtp(2).Text, PathFileIni
WriteINI "ftpconnect", "passobl", txtFtp(3).Text, PathFileIni
'===========

'ФТП київський
WriteINI "ftpconnect", "ftpkiev", txtFtp(4).Text, PathFileIni
WriteINI "ftpconnect", "kievkat", txtFtp(5).Text, PathFileIni
WriteINI "ftpconnect", "logkiev", txtFtp(6).Text, PathFileIni
WriteINI "ftpconnect", "passkiev", txtFtp(7).Text, PathFileIni
'===========
cmbSaveFtp.Enabled = False
End Sub

Private Sub cmbSaveSR_Click()
WriteINI "SendReport", "Pathrpr", txtSR(0).Text, PathFileIni
WriteINI "SendReport", "MailFolder", txtSR(1).Text, PathFileIni
'1-й операційний
If Chk_Oper1.Value = 1 Then
    WriteINI "SendReport", "Oper1", "1", PathFileIni
Else
    WriteINI "SendReport", "Oper1", "0", PathFileIni
End If
'2-й операційний
If Chk_Oper2.Value = 1 Then
    WriteINI "SendReport", "Oper2", "1", PathFileIni
Else
    WriteINI "SendReport", "Oper2", "0", PathFileIni
End If
'3-й операційний
If Chk_Oper3.Value = 1 Then
    WriteINI "SendReport", "Oper3", "1", PathFileIni
Else
    WriteINI "SendReport", "Oper3", "0", PathFileIni
End If
'4-й операційний
If Chk_Oper4.Value = 1 Then
    WriteINI "SendReport", "Oper4", "1", PathFileIni
Else
    WriteINI "SendReport", "Oper4", "0", PathFileIni
End If

cmbSaveSR.Enabled = False
End Sub

Private Sub cmbSaveVP_Click()
WriteINI "viezd", "PathSend", txtVP(0).Text, PathFileIni
'WriteINI "viezd", "pos", txtVP(1).Text, PathFileIni
WriteINI "viezd", "tel", txtVP(2).Text, PathFileIni
WriteINI "viezd", "prf", txtVP(3).Text, PathFileIni
WriteINI "viezd", "ID", txtVP(4).Text, PathFileIni
cmbSaveVP.Enabled = False
End Sub

Private Sub cmbSelFolder_SR1_Click()
txtSR(0).Text = BrowseForFolder(Me.hwnd, "Оберіть каталог районних БД." & _
                vbCrLf & "Зазвичай це V:\", "V:\")
cmbSaveSR.Enabled = True
End Sub

Private Sub cmbSelFolder_SR2_Click()
txtSR(1).Text = BrowseForFolder(Me.hwnd, "Оберіть поштовий каталог." & _
                vbCrLf & "Зазвичай це S:\", "S:\")
cmbSaveSR.Enabled = True
End Sub

Private Sub cmbSelFolder_VP_Click()
txtVP(0).Text = BrowseForFolder(Me.hwnd, "Оберіть каталог передачі", "c:\send\out\")
cmbSaveVP.Enabled = True
End Sub

Private Sub cmbSelFolderBD_Click()
On Error GoTo ErrorHandler
With CommonDialog1
    .Flags = cdlOFNHideReadOnly
    .Filter = "База даних (DB.mdb)|DB.mdb| Усі файли (*.*)|*.*"
    .InitDir = App.Path
    .DialogTitle = "Оберіть базу данних."
    .ShowOpen
    txtBD.Text = .FileName
End With
ErrorHandler:
If Err.Number = 32755 Then
   Exit Sub
End If

cmbSaveBD.Enabled = True
End Sub

Private Sub cmbTEST_Click(Index As Integer)

hOpen = InternetOpen("My VB Test", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
  If hOpen = 0 Then
    ErrorOut Err.LastDllError, "InternetOpen"
    Unload Me
  End If
  dwType = FTP_TRANSFER_TYPE_ASCII
  dwSeman = 0
  hConnection = 0

Select Case Index
    'тест з'єднання Обласного ФТП
    Case 0
      If hConnection <> 0 Then
        InternetCloseHandle hConnection
    End If
    hConnection = InternetConnect(hOpen, txtFtp(0), INTERNET_INVALID_PORT_NUMBER, _
    txtFtp(2), txtFtp(3), INTERNET_SERVICE_FTP, dwSeman, 0)
    If hConnection = 0 Then
        ErrorOut Err.LastDllError, "Connect with " & txtFtp(0).Text
        Exit Sub
    Else
        MsgBox "Підключено!", , txtFtp(0)
    End If
    '===========
    
    'тест з'єднання Київського ФТП
    Case 1
      If hConnection <> 0 Then
        InternetCloseHandle hConnection
    End If
    hConnection = InternetConnect(hOpen, txtFtp(4), INTERNET_INVALID_PORT_NUMBER, _
    txtFtp(6), txtFtp(7), INTERNET_SERVICE_FTP, dwSeman, 0)
    If hConnection = 0 Then
        ErrorOut Err.LastDllError, "Connect with " & txtFtp(0).Text
        Exit Sub
    Else
        MsgBox "Підключено!", , txtFtp(4)
    End If
End Select
End Sub

Private Sub Form_Load()
SSTab1.TabVisible(0) = False
frmMain.Setting.Enabled = False
Me.Width = 11000
Me.Height = 8500

' виїзд пенсіонера
txtVP(0).Text = ReadINI("viezd", "PathSend", PathFileIni)
'txtVP(1).Text = ReadINI("viezd", "pos", PathFileIni)
txtVP(2).Text = ReadINI("viezd", "tel", PathFileIni)
txtVP(3).Text = ReadINI("viezd", "prf", PathFileIni)
txtVP(4).Text = ReadINI("viezd", "ID", PathFileIni)
'=================

'FTP обласний
txtFtp(0).Text = ReadINI("ftpconnect", "ftpobl", PathFileIni)
'каталог
txtFtp(1).Text = ReadINI("ftpconnect", "oblkat", PathFileIni)
'логін
txtFtp(2).Text = ReadINI("ftpconnect", "logobl", PathFileIni)
'пароль
txtFtp(3).Text = ReadINI("ftpconnect", "passobl", PathFileIni)
'================

'FTP київський
txtFtp(4).Text = ReadINI("ftpconnect", "ftpkiev", PathFileIni)
'каталог
txtFtp(5).Text = ReadINI("ftpconnect", "kievkat", PathFileIni)
'логін
txtFtp(6).Text = ReadINI("ftpconnect", "logkiev", PathFileIni)
'пароль
txtFtp(7).Text = ReadINI("ftpconnect", "passkiev", PathFileIni)
'===============

'надіслати звіти
txtSR(0).Text = ReadINI("SendReport", "PathRpr", PathFileIni)
txtSR(1).Text = ReadINI("SendReport", "MailFolder", PathFileIni)
'===============

'основні
txtBD.Text = ReadINI("DataBase", "dbpath", PathFileIni)
'==============

'райони по операційних
If ReadINI("SendReport", "Oper1", PathFileIni) = 1 Then Chk_Oper1.Value = 1
If ReadINI("SendReport", "Oper2", PathFileIni) = 1 Then Chk_Oper2.Value = 1
If ReadINI("SendReport", "Oper3", PathFileIni) = 1 Then Chk_Oper3.Value = 1
If ReadINI("SendReport", "Oper4", PathFileIni) = 1 Then Chk_Oper4.Value = 1
'==============

'кнопки зберегти неактивні
cmbSaveVP.Enabled = False
cmbSaveSR.Enabled = False
cmbSaveBD.Enabled = False
cmbSaveFtp.Enabled = False
'===============
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Setting.Enabled = True
InternetCloseHandle hOpen
Unload Me
End Sub

Private Sub txtBD_KeyPress(KeyAscii As Integer)
If KeyAscii <> 0 Then cmbSaveBD.Enabled = True
End Sub

Private Sub txtFtp_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 0 Then cmbSaveFtp.Enabled = True
End Sub

Private Sub txtFtp_Validate(Index As Integer, Cancel As Boolean)
If Right(txtFtp(1).Text, 1) <> "/" Then txtFtp(1).Text = txtFtp(1).Text & "/"
If Right(txtFtp(5).Text, 1) <> "/" Then txtFtp(5).Text = txtFtp(5).Text & "/"

If Left(txtFtp(1).Text, 1) <> "/" Then txtFtp(1).Text = "/" & txtFtp(1).Text
If Left(txtFtp(5).Text, 1) <> "/" Then txtFtp(5).Text = "/" & txtFtp(5).Text
End Sub

Private Sub txtSR_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 0 Then cmbSaveSR.Enabled = True
End Sub

Private Sub txtVP_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 0 Then cmbSaveVP.Enabled = True
End Sub

Private Sub ErrorOut(ByVal dwError As Long, ByRef szFunc As String)
Dim dwRet As Long
Dim dwTemp As Long
Dim szString As String * 2048
Dim szErrorMessage As String

dwRet = FormatMessage(FORMAT_MESSAGE_FROM_HMODULE, _
                  GetModuleHandle("wininet.dll"), dwError, 0, _
                  szString, 256, 0)
szErrorMessage = szFunc & vbCrLf & " Код помилки: " & dwError & vbCrLf & " Текст помилки: " & szString
Debug.Print szErrorMessage
MsgBox szErrorMessage, , "Увага"
If (dwError = 12003) Then
    ' Extended error information was returned
    dwRet = InternetGetLastResponseInfo(dwTemp, szString, 2048)
    Debug.Print szString
    frmErr.Show
    frmErr.Text1.Text = szString
End If
End Sub

Private Sub txtVP_Validate(Index As Integer, Cancel As Boolean)
If Right(txtVP(0).Text, 1) <> "\" Then txtVP(0).Text = txtVP(0).Text & "\"
End Sub
