VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmVP 
   Caption         =   "Виїзд пенсіонера"
   ClientHeight    =   7200
   ClientLeft      =   5100
   ClientTop       =   3630
   ClientWidth     =   9075
   DrawStyle       =   2  'Dot
   Icon            =   "frmVP.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7450.111
   ScaleMode       =   0  'User
   ScaleWidth      =   10383.29
   Begin VB.CommandButton cmbSendFtp 
      Caption         =   "Відправити"
      Height          =   387
      Left            =   240
      TabIndex        =   27
      ToolTipText     =   "Відправити архів на FTP-сервер"
      Top             =   5040
      Width           =   2025
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H8000000F&
      Height          =   2175
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   3840
      Width           =   6375
   End
   Begin VB.CommandButton cmbArh 
      Caption         =   "Архівувати"
      Height          =   387
      Left            =   240
      TabIndex        =   12
      ToolTipText     =   "Архівація супровідного і файла ел.пенсіної справи"
      Top             =   4440
      Width           =   2025
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      DragMode        =   1  'Automatic
      Height          =   345
      Left            =   0
      TabIndex        =   25
      Top             =   6855
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   2778
            MinWidth        =   2774
            Text            =   "Каталог передачі:"
            TextSave        =   "Каталог передачі:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12674
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Дані пенсіонера"
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
      Height          =   1215
      Index           =   0
      Left            =   240
      TabIndex        =   21
      Top             =   2400
      Width           =   8535
      Begin VB.TextBox txtFile 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   720
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   8
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtPib 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   4080
         TabIndex        =   9
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox txtOsob 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   720
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label19 
         Caption         =   "Файл"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label18 
         Caption         =   "ПІБ пенсіонера"
         Height          =   315
         Left            =   2760
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "ОР №"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdExportWord 
      Caption         =   "Супровідний лист"
      Height          =   387
      Left            =   240
      TabIndex        =   11
      ToolTipText     =   "Сформувати супровідний лист"
      Top             =   3840
      Width           =   2025
   End
   Begin VB.Frame Frame1 
      Caption         =   "Область та район вибуття"
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
      Index           =   2
      Left            =   240
      TabIndex        =   18
      Top             =   840
      Width           =   2895
      Begin VB.TextBox txtRajFrom 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   4
         Top             =   720
         Width           =   1080
      End
      Begin VB.TextBox txtOblFrom 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   3
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label14 
         Caption         =   "Район"
         Height          =   225
         Left            =   480
         TabIndex        =   20
         Top             =   750
         Width           =   570
      End
      Begin VB.Label Label13 
         Caption         =   "Область"
         Height          =   225
         Left            =   480
         TabIndex        =   19
         Top             =   375
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Область та район прибуття"
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
      Index           =   3
      Left            =   3360
      TabIndex        =   16
      Top             =   840
      Width           =   5415
      Begin VB.ComboBox txtRajTo 
         Height          =   315
         Left            =   720
         TabIndex        =   6
         Top             =   720
         Width           =   4455
      End
      Begin VB.ComboBox txtOblTo 
         Height          =   315
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label lblNameObl 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblNameRaj 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   465
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Дані листа-запиту на висилку"
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
      Height          =   975
      Index           =   1
      Left            =   5880
      TabIndex        =   13
      Top             =   6000
      Visible         =   0   'False
      Width           =   2895
      Begin MSMask.MaskEdBox dataF 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "№"
         Height          =   195
         Left            =   720
         TabIndex        =   15
         Top             =   600
         Width           =   195
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Запит від "
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.CommandButton cmbImport 
      BackColor       =   &H000000C0&
      Caption         =   "Імпорт 1LS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6960
      TabIndex        =   2
      Top             =   240
      Width           =   1770
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmVP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public myFrmMain As frmMain

Public NameArhive As String      ' назва архіва 1ls + doc
Public NameDoc As String         ' назва супровідної
Public PathKatalogSend As String ' каталог передачі
   
Dim hConnection As Long, _
          hOpen As Long, _
          hFile As Long, _
         dwType As Long, _
        dwSeman As Long
Private cCombo As New clsAutoCombo
Private Declare Function ActivateKeyboardLayout& Lib "user32" (ByVal HKL As Long, _
ByVal Flags As Long)


Private Sub cmbArh_Click()
Dim myCmd As String, tmpNameFile, Quest
      
tmpNameFile = Split(txtFile.Text, ".")
NameArhive = txtOblFrom & txtRajFrom & "_to_" & _
                lblNameObl & lblNameRaj & "_" & Right(tmpNameFile(0), 4) & ".rar"

'перевірка наявності вигрузки 1LS
If File_Exists(PathKatalogSend & txtFile) = False Then
    MsgBox "Файл вигрузки відсутній", vbCritical, "Архівація не можлива!"
'перевірка наявності супровідної
ElseIf File_Exists(PathKatalogSend & Left(txtFile, 8) & ".doc") = False Then
    Quest = MsgBox(PathKatalogSend & Left(txtFile, 8) & ".doc" & vbCrLf & _
    "Архівувати електронну пенсійну справу без супровідної?", vbExclamation + vbYesNo, "Файл супровідної відсутній!")
    If Quest = vbYes Then
        Call Shell("C:\Program Files\WinRAR\winrar.exe a -ep " & PathKatalogSend & _
        NameArhive & " " & PathKatalogSend & txtFile, vbNormalFocus)
        With txtLog
            .Text = .Text & "=======" & vbCrLf & Time & " @ Архівація пройшла БЕЗ супровідного листа. Завершено УСПІШНО!" & vbCrLf & _
            Time & " = Створено архів " & PathKatalogSend & NameArhive & vbCrLf & "=======" & vbCrLf
        End With
    End If
    
    If Quest = vbNo Then
        txtLog.Text = txtLog.Text & Time & " @ Натисніть кнопку <Супровідний лист>" & vbCrLf
    End If
Else
    Call Shell("C:\Program Files\WinRAR\winrar.exe a -ep " & PathKatalogSend & _
    NameArhive & " " & PathKatalogSend & tmpNameFile(0) & ".*", vbNormalFocus)
    Sleep 500
    If File_Exists(PathKatalogSend & NameArhive) = True Then
    txtLog.Text = txtLog.Text & "=======" & vbCrLf & Time & " = Архівація пройшла успішно!" & vbCrLf & _
            Time & " = Створено архів " & PathKatalogSend & NameArhive & vbCrLf & "=======" & vbCrLf
    DoEvents
    Sleep 2500
    cmbSendFtp.Enabled = True
    cmbSendFtp.SetFocus
    End If
            
End If

End Sub

Private Sub cmbImport_Click()
Dim f, s, fio, fio1, FL, Sel_Obl
Dim v As Variant
Dim SS As String, sLS As String
Dim S1 As String, s2 As String
Dim I As Integer, j As Integer
Dim otvet As Integer

On Error GoTo ErrorHandler
With CommonDialog1
    .Flags = cdlOFNHideReadOnly
    .Filter = "Файли 1LS (*.1ls)|*.1ls| Усі файли (*.*)|*.*"
    .InitDir = PathKatalogSend
    .DialogTitle = "Оберіть вигрузку електронної пенсійної справи."
    .ShowOpen
End With

FL = CommonDialog1.FileName
If FL = "" Or IsNull(FL) = True Then End

f = FreeFile
Open FL For Input As #f
SS = ""
sLS = ""

txtLog.Text = Time & " = Відкрито файл " & FL & vbCrLf

Do Until EOF(f)
Line Input #f, s
If Left(s, 3) = "$S," Then SS = s
If Left(s, 3) = "$LS" Then sLS = s
Loop
Close #f

If SS > "" Then
For Each v In Split(SS, ",")
Select Case Left(v, 2)
Case "B="
    txtRajFrom = Mid(v, 5, 2)
Case "F="
    fio = Mid(v, 4, Len(v) - 4)
    fio1 = Dos2Win(fio)
  For I = 1 To 5
    Select Case I
    Case 1: S1 = Chr(175)
            s2 = "Є"
    Case 2: S1 = Chr(161)
            s2 = "і"
    Case 3: S1 = Chr(176)
            s2 = "ї"
    Case 4: S1 = Chr(34)
            s2 = "`"
    Case 5: S1 = Chr(39)
            s2 = "`"
    End Select
    j = InStr(1, fio1, S1)
    While j <> 0
      fio1 = Left(fio1, j - 1) + s2 + Mid(fio1, j + 1)
      j = InStr(j + 1, fio1, S1)
    Wend
  Next I
txtPib = StrConv(fio1, vbProperCase)
End Select

Next
End If

    If sLS > "" Then
    txtOsob = Val(Mid(sLS, 5))
End If
txtFile = Dir(FL)
'текст в лог
txtLog.Text = txtLog.Text & Time & " = Дані успішно імпортовані." & vbCrLf

otvet = MsgBox("Виїзд в межах області??", vbYesNo + vbInformation, "УВАГА")
Select Case otvet
Case vbYes
    txtOblFrom.Text = ReadINI("Viezd", "KodObl", PathFileIni)
    lblNameObl.Caption = txtOblFrom.Text
    'added
    'sql to get obl name
    
    ConnectToDataBase
Sel_Obl = "SELECT obl_name FROM obl WHERE obl_kod =" & txtOblFrom.Text
myRS.Open Sel_Obl, myADO, adOpenStatic
Do While Not myRS.EOF
    txtOblTo.Text = myRS("obl_name").Value
    myRS.MoveNext
Loop
    'added
    'With txtOblTo
    '    .Text = "Хмельницькій області"
    '    .SetFocus
    'End With
    'txtRajTo.SetFocus
Case vbNo
    txtOblFrom = ReadINI("Viezd", "KodOblUa", PathFileIni)
    txtOblTo.SetFocus
End Select

ErrorHandler:
If Err.Number = 32755 Then
   Exit Sub
End If

End Sub


Private Sub cmbSendFtp_Click()
Dim tmpSql, tmpSQL1 As String
Dim sTmpFtp, _
    sTmpKat, _
    sTmpLog, _
    stmpPass, _
    sTmpRajOrObl

' якщо код області локальний
If txtOblFrom = ReadINI("Viezd", "KodObl", PathFileIni) Then
    sTmpFtp = Ftp_Obl
    sTmpKat = Kat_Obl
    sTmpLog = Login_Obl
    stmpPass = Pass_Obl
    sTmpRajOrObl = lblNameRaj & "/"
End If
' якщо код області по УКРАЇНІ
If txtOblFrom = ReadINI("Viezd", "KodOblUa", PathFileIni) Then
    sTmpFtp = Ftp_Kiev
    sTmpKat = Kat_Kiev
    sTmpLog = Login_Kiev
    stmpPass = Pass_Kiev
    sTmpRajOrObl = lblNameObl & "/"
End If

' зєднуємось з фтп ------------------------------------------------------
    If hConnection <> 0 Then
        InternetCloseHandle hConnection
    End If
    hConnection = InternetConnect(hOpen, sTmpFtp, INTERNET_INVALID_PORT_NUMBER, _
    sTmpLog, stmpPass, INTERNET_SERVICE_FTP, dwSeman, 0)
    If hConnection = 0 Then
        ErrorOut Err.LastDllError, "InternetConnect"
        Exit Sub
    Else
        txtLog.Text = txtLog.Text & Time & " = Встановлено з'єднання з " & sTmpFtp & vbCrLf
    End If
    
'==========================================================================
   Sleep 300
' встановлюємо цільовий каталог -----------------------------------------
    If (FtpSetCurrentDirectory(hConnection, sTmpKat & sTmpRajOrObl) = False) Then
        ErrorOut Err.LastDllError, "FtpSetCurrentDirectory"
        Exit Sub
    Else
        txtLog.Text = txtLog.Text & Time & " = Каталог для копіювання " & sTmpKat & sTmpRajOrObl & vbCrLf
    End If
'==========================================================================
Sleep 300
'кидаєм файл в папку ----------------------------------
    If (FtpPutFile(hConnection, PathKatalogSend & NameArhive, sTmpKat & sTmpRajOrObl & NameArhive, _
        dwType, 0) = False) Then
        ErrorOut Err.LastDllError, "FtpPutFile"
        Exit Sub
    Else
        txtLog.Text = txtLog.Text & Time & " = " & NameArhive & " скопійовано в " & sTmpKat & sTmpRajOrObl & vbCrLf
            
    End If
'==========================================

'перевірка на наявність каталога РІК_МІСЯЦЬ

    If Folder_Exists(PathKatalogSend & Folder_God_Mes) = False Then
    MkDir (PathKatalogSend & Folder_God_Mes)
    MkDir (PathKatalogSend & Folder_God_Mes & "\" & "Arhiv")
    Else
        If Folder_Exists(PathKatalogSend & Folder_God_Mes & "\" & "Arhiv") = False Then
        MkDir (PathKatalogSend & Folder_God_Mes & "\" & "Arhiv")
        End If
    End If
'==========================================
    txtLog.Text = txtLog.Text & "=======" & vbCrLf
    'перевірка на наявність вигрузки 1лс
    If File_Exists(PathKatalogSend & txtFile.Text) = False Then
    txtLog.Text = txtLog.Text & Time & " @ Відсутній файл вигрузки 1LS " & PathKatalogSend & txtFile.Text & vbCrLf
    Else
    'перенесення вигрузки 1лс в каталог РІК_МІСЯЦЬ
    FileCopy PathKatalogSend & txtFile.Text, PathKatalogSend & Folder_God_Mes & "\" & txtFile.Text
    Kill (PathKatalogSend & txtFile.Text)
    txtLog.Text = txtLog.Text & Time & " = " & txtFile.Text & " перенесено в " & PathKatalogSend & Folder_God_Mes & "\" & vbCrLf
    End If
    '======================================
    Sleep 300
    'перевірка на наявність  супровідної
    If File_Exists(PathKatalogSend & NameDoc) = False Then
    txtLog.Text = txtLog.Text & Time & " @ Відсутня супровідна " & PathKatalogSend & NameDoc & vbCrLf
    Else
    'перенесення супровідної в каталог РІК_МІСЯЦЬ
    FileCopy PathKatalogSend & NameDoc, PathKatalogSend & Folder_God_Mes & "\" & NameDoc
    Kill (PathKatalogSend & NameDoc)
    txtLog.Text = txtLog.Text & Time & " = " & NameDoc & " перенесено в " & PathKatalogSend & Folder_God_Mes & "\" & vbCrLf
    End If
    '======================================
    Sleep 300
    'перевірка на наявність АРХІВА
    If File_Exists(PathKatalogSend & NameArhive) = False Then
    txtLog.Text = txtLog.Text & Time & " @ Відсутній архів " & PathKatalogSend & NameArhive & vbCrLf
    Else
    'перенесення АРХІВАв каталог РІК_МІСЯЦЬ\arhiv
    FileCopy PathKatalogSend & NameArhive, PathKatalogSend & Folder_God_Mes & "\" & "Arhiv" & "\" & NameArhive
    Kill (PathKatalogSend & NameArhive)
    txtLog.Text = txtLog.Text & Time & " = " & NameArhive & " перенесено в " & PathKatalogSend & Folder_God_Mes & "\" & "Arhiv" & vbCrLf
    End If
    
'Set myADO = Nothing
'Set myRS = Nothing
ConnectToDataBase

tmpSql = "INSERT INTO DEP (dep_or, dep_date, dep_ins,  dep_pib, dep_from, dep_to, dep_nfl) " & _
        "VALUES (" & _
        txtOsob.Text & ",'" & _
        Date & Chr(32) & Time() & "'," & _
        ReadINI("viezd", "ID", PathFileIni) & ",'" & _
        txtPib.Text & "'," & txtOblFrom.Text & _
        txtRajFrom.Text & ",'" & _
        lblNameObl & lblNameRaj & "','" & _
        txtFile.Text & "')"
'Text2.Text = tmpSql
myRS.Open tmpSql, myADO, adOpenDynamic
 txtLog.Text = txtLog.Text & Time & " = Запис додано в базу" & vbCrLf & _
    txtLog.Text & Time & " = Обробку виїзду завершено"
MsgBox "Обробку виїзду завершено", vbInformation, "Усё здєлано, шеф"
End Sub

Private Sub cmdExportWord_Click()
Dim appWord As Word.Application
Dim docWord As Word.Document
Dim rngCurrent As Word.Range
Dim objTable As Word.Table
Dim s As String
Dim iTBL_Rows As Integer
   
On Error GoTo Err_AccessToWord

    Set appWord = CreateObject("Word.Application")
    
    With appWord
             .Visible = False
             '.Activate
             .WindowState = wdWindowStateNormal
         End With
         
    'Створюєм документ
        Set docWord = appWord.Documents.Add
         
        With docWord.PageSetup
            'альбомна орієнтація
            .Orientation = wdOrientLandscape
            
            ' відступи колонтитулів
            .HeaderDistance = appWord.CentimetersToPoints(0.5)
            .FooterDistance = appWord.CentimetersToPoints(1)
            
            'відступи на сторінці
            .TopMargin = appWord.CentimetersToPoints(1)
            .LeftMargin = appWord.CentimetersToPoints(2)
            .BottomMargin = appWord.CentimetersToPoints(1)
            .RightMargin = appWord.CentimetersToPoints(1)
            
            'розмір сторінки
            .PageWidth = appWord.CentimetersToPoints(16)
            .PageHeight = appWord.CentimetersToPoints(15)
        End With
        
    '--------------------------------------------
    'колонтітул нижній - дата
With docWord.Sections(1)
    .Footers(wdHeaderFooterPrimary).Range.Text = Date
    '.Footers(wdHeaderFooterPrimary).PageNumbers.Add
    .Footers(wdHeaderFooterPrimary).Range.Select
    appWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    appWord.Selection.Font.Name = "Times New Roman"
    appWord.Selection.Font.Size = 8
End With
    
With docWord.ActiveWindow
    .ActivePane.close
    .View = wdPrintView
End With
          
    'пишем ПРОТОКОЛ
    Set rngCurrent = appWord.ActiveDocument.Sections(1).Range
    With rngCurrent
        '.InsertParagraphAfter
        .Collapse Direction:=wdCollapseEnd
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Text = "ПРОТОКОЛ ПЕРЕДАЧІ ЕЛЕКТРОННОЇ ПЕНСІЙНОЇ СПРАВИ"
        .Select
        .Font.Name = "Times New Roman"
        .Font.Size = 13
        .Font.Bold = False
    End With
        
Set rngCurrent = appWord.ActiveDocument.Sections(1).Range
With rngCurrent
    .InsertParagraphAfter
    .Collapse Direction:=wdCollapseEnd
End With
        
    'табличка 2 на 5
    Set rngCurrent = appWord.ActiveDocument.Sections(1).Range
    With rngCurrent
        .InsertParagraphAfter
        .Collapse Direction:=wdCollapseEnd
    End With
        
    Set objTable = docWord.Tables.Add(Range:=rngCurrent, NumRows:=iTBL_Rows + 5, NumColumns:=2)
    'added
    
    
    Dim Sel_Obl As String
    
    ConnectToDataBase
Sel_Obl = "SELECT obl_name , [Raj_Name] " & _
            "From OBL, RAJ " & _
            "Where OBL.obl_kod = RAJ.Raj_Obl And obl_kod = " & txtOblFrom.Text & " And RAJ.Raj_Kod = " & txtRajFrom.Text
myRS.Open Sel_Obl, myADO, adOpenStatic


Dim FullNameObl As String
Dim FullNameRaj As String

If (myRS.RecordCount > 0) Then
FullNameObl = myRS("obl_name").Value
FullNameRaj = myRS("Raj_Name").Value
End If

    'added
    With objTable
        .Borders.Enable = False
        .Rows.Height = 10
        .Columns.Width = 50
        'називаєм колонки
        .Cell(2, 1).Range.Text = "Область та " & vbCrLf & "район отримувач:"
        .Cell(2, 2).Range.Text = txtRajTo.Text & " (" & lblNameObl & lblNameRaj & ")"
        
        .Cell(1, 1).Range.Text = "Область та " & vbCrLf & "район вибуття:"
        .Cell(1, 2).Range.Text = FullNameObl & " " & FullNameRaj
        '"Хмельницька обл.," & vbCrLf & Select_Raj_Hm(txtRajFrom.Text) & _
                                "  (" & txtOblFrom.Text & txtRajFrom.Text & ")"
        
        .Cell(3, 1).Range.Text = "№ ел. пенсійної справи:"
        .Cell(3, 2).Range.Text = txtOsob.Text
        
        .Cell(4, 1).Range.Text = "ПІБ:"
        .Cell(4, 2).Range.Text = txtPib.Text
        
        .Cell(5, 1).Range.Text = "Ім'я файлу: "
        .Cell(5, 2).Range.Text = txtFile.Text
    End With

        objTable.Select
        appWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        appWord.Selection.Font.Size = 14
        appWord.Selection.MoveDown
       
        With objTable
        
        With .Borders(wdBorderHorizontal)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderVertical)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        
        .AllowAutoFit = True
            .Columns(1).Width = 120
            .Columns(2).Width = 220
            .Columns.Select
            appWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Columns(1).Select
            appWord.Selection.Font.Bold = False
            appWord.Selection.Font.Size = 13
            appWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Columns(2).Select
            appWord.Selection.Font.Bold = True
            appWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        End With
        
        Set rngCurrent = appWord.ActiveDocument.Sections(1).Range
        With rngCurrent
            '.InsertParagraphAfter
            .Collapse Direction:=wdCollapseEnd
        End With
                

        Set objTable = Nothing

'пишем протокол з БУФЕРА ОБМІНУ
    Set rngCurrent = appWord.ActiveDocument.Sections(1).Range
    With rngCurrent
    .InsertParagraphAfter
    .Collapse Direction:=wdCollapseEnd
    .ParagraphFormat.Alignment = wdAlignParagraphLeft
    .Text = Clipboard.GetText
    .Select
    .Font.Name = "Lucida Console"
    .Font.Size = 8
    .Font.Bold = False
    .InsertParagraphAfter
    '.InsertParagraphAfter
    End With

'виконавець
    Set rngCurrent = appWord.ActiveDocument.Sections(1).Range
        
    With rngCurrent
    '.InsertParagraphBefore
    .Collapse Direction:=wdCollapseEnd
    .ParagraphFormat.Alignment = wdAlignParagraphJustify
    .Text = "Відповідальна особа: " & vbCrLf & frmMain.StatusBar1.Panels(8).Text & _
            vbCrLf & "тел. " & ReadINI("viezd", "tel", PathFileIni) ' ПІДПИС ВИКОНАВЕЦЬ
    '.InsertParagraphAfter
    .Select
    .Font.Name = "Times New Roman"
    .Font.Size = 12
    '.InsertParagraphAfter
    End With
    
        NameDoc = Left(txtFile.Text, 8) & ".doc"
        'appWord.Visible = True
         
        docWord.SaveAs PathKatalogSend & NameDoc 'ЗБЕРІГАЄ ФАЙЛ
    
    'Set objTable = Nothing
    Set rngCurrent = Nothing
    Set docWord = Nothing 'документ
    appWord.Quit
    Set appWord = Nothing ' приложение
    txtLog.Text = txtLog.Text & Time & " = Створено супровідну " & PathKatalogSend & NameDoc & vbCrLf
   
   
   'cmdExportWord.Enabled = False
    cmbArh.Enabled = True

Err_cmdExportWord_Click:
Exit Sub

Err_AccessToWord:
    MsgBox "The Following Automation Error has occurred:" & vbCrLf & Err.Description & "  " & Err.Number, vbCritical, "Automation Error!"
    Resume Err_cmdExportWord_Click
    

End Sub


Private Sub dataF_KeyPress(KeyAscii As Integer)
TextPressEnter (KeyAscii)
End Sub

Private Sub Form_Load()
Dim Sel_Obl
Me.Width = 9200
Me.Height = 8000

'dataF.Text = Date
cmdExportWord.Enabled = False
cmbArh.Enabled = False
cmbSendFtp.Enabled = False

PathKatalogSend = ReadINI("viezd", "pathsend", PathFileIni)

With StatusBar1.Panels(2)
    .Text = PathKatalogSend
    .ToolTipText = "Каталог передачі"
End With

' фтп зєднання
  hOpen = InternetOpen("My VB Test", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
  If hOpen = 0 Then
    ErrorOut Err.LastDllError, "InternetOpen"
    'Unload Form1
  End If
  dwType = FTP_TRANSFER_TYPE_ASCII
  dwSeman = 0
  hConnection = 0
' ===========
ConnectToDataBase
Sel_Obl = "SELECT obl_kod, obl_name FROM obl ORDER BY obl_name"
myRS.Open Sel_Obl, myADO, adOpenStatic
Do While Not myRS.EOF
    With txtOblTo
    .AddItem myRS("obl_name").Value
    .ItemData(txtOblTo.NewIndex) = Val(myRS("obl_kod"))
    myRS.MoveNext
    End With
Loop
ActivateKeyboardLayout &H4220422, 3
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
InternetCloseHandle hOpen

myFrmMain.viezd.Enabled = True
Set myRS = Nothing
Set myADO = Nothing
End Sub

Private Sub txtFile_KeyPress(KeyAscii As Integer)
TextPressEnter (KeyAscii)
End Sub

Private Sub txtFile_Validate(Cancel As Boolean)
If Len(txtFile.Text) > 0 Then
    cmdExportWord.Enabled = True
    ElseIf Len(txtFile.Text) = 0 Then cmdExportWord.Enabled = False And _
                                             cmbArh.Enabled = False
End If
End Sub

Private Sub txtLog_Change()
txtLog.SelStart = Len(txtLog)
End Sub

Private Sub txtOblTo_Click()
With txtOblTo
lblNameObl.Caption = CStr(.ItemData(.ListIndex))
End With
If Len(lblNameObl.Caption) < 2 Then lblNameObl.Caption = "0" & lblNameObl.Caption
End Sub

Private Sub txtOblTo_KeyPress(KeyAscii As Integer)
KeyAscii = cCombo.AutoFind(txtOblTo, KeyAscii, True)
End Sub

Private Sub txtOblTo_Validate(Cancel As Boolean)
Dim Sel_Raj
Set myRS = Nothing
Sel_Raj = "SELECT * FROM raj WHERE raj_obl = '" & lblNameObl & Chr(39) & "ORDER BY raj_name"
txtRajTo.Clear
ConnectToDataBase
myRS.Open Sel_Raj, myADO, adOpenStatic
Do While Not myRS.EOF
With txtRajTo
    .AddItem myRS("raj_name").Value
    .ItemData(.NewIndex) = Val(myRS("raj_kod"))
    myRS.MoveNext
    End With
Loop
        
End Sub

Private Sub txtOsob_KeyDown(KeyCode As Integer, Shift As Integer)
With txtOsob
      .Locked = IIf((KeyCode > 47 And KeyCode < 58) Or _
      (KeyCode > 95 And KeyCode < 107) Or _
      (KeyCode = 8) Or (KeyCode = 46) Or (KeyCode = 188), IIf(KeyCode = 188, _
      IIf(InStr(1, Text1, ",") = 0 And .SelStart <> 0, False, True), False), True)
End With
End Sub


Private Sub txtRajFrom_KeyDown(KeyCode As Integer, Shift As Integer)
With txtRajFrom
      .Locked = IIf((KeyCode > 47 And KeyCode < 58) Or _
      (KeyCode > 95 And KeyCode < 107) Or _
      (KeyCode = 8) Or (KeyCode = 46) Or (KeyCode = 188), IIf(KeyCode = 188, _
      IIf(InStr(1, Text1, ",") = 0 And .SelStart <> 0, False, True), False), True)
End With

End Sub
Private Sub ErrorOut(ByVal dwError As Long, ByRef szFunc As String)
Dim dwRet As Long
Dim dwTemp As Long
Dim szString As String * 2048
Dim szErrorMessage As String

dwRet = FormatMessage(FORMAT_MESSAGE_FROM_HMODULE, _
                  GetModuleHandle("wininet.dll"), dwError, 0, _
                  szString, 256, 0)
szErrorMessage = szFunc & " error code: " & dwError & " Message: " & szString
Debug.Print szErrorMessage
MsgBox szErrorMessage, , "Ftp"
If (dwError = 12003) Then
    ' Extended error information was returned
    dwRet = InternetGetLastResponseInfo(dwTemp, szString, 2048)
    Debug.Print szString
    txtLog.Text = szString
End If
End Sub


Private Function ConnectToFtp(ServName As String, UserName As String, PassWord As String)
    If hConnection <> 0 Then
        InternetCloseHandle hConnection
    End If
    hConnection = InternetConnect(hOpen, ServName, INTERNET_INVALID_PORT_NUMBER, _
    UserName, PassWord, INTERNET_SERVICE_FTP, dwSeman, 0)
    If hConnection = 0 Then
        ErrorOut Err.LastDllError, "InternetConnect"
        Exit Function
    Else
    ConnectToFtp = "Connect to " & ServName
End If
End Function

Public Function Folder_God_Mes()
Folder_God_Mes = Year(Date) & "_" & Month(Date)
End Function


Private Sub txtRajTo_Click()
With txtRajTo
lblNameRaj.Caption = CStr(.ItemData(.ListIndex))
End With
With lblNameRaj
    If Len(.Caption) < 2 Then
    .Caption = "0" & .Caption
    End If
End With
cmdExportWord.Enabled = True

End Sub

Private Sub txtRajTo_KeyPress(KeyAscii As Integer)
KeyAscii = cCombo.AutoFind(txtRajTo, KeyAscii, True)
End Sub

Private Sub txtRajTo_Validate(Cancel As Boolean)
cmdExportWord.SetFocus
End Sub
