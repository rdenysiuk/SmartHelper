VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmVP 
   Caption         =   "Виїзд пенсіонера"
   ClientHeight    =   7485
   ClientLeft      =   5100
   ClientTop       =   3630
   ClientWidth     =   9075
   DrawStyle       =   2  'Dot
   Icon            =   "frmVP.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7745.011
   ScaleMode       =   0  'User
   ScaleWidth      =   10383.29
   Begin VB.CommandButton cmbSendFtp 
      Caption         =   "Відправити"
      Height          =   387
      Left            =   240
      TabIndex        =   29
      ToolTipText     =   "Відправити архів на FTP-сервер"
      Top             =   6000
      Width           =   2025
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H8000000F&
      Height          =   1575
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      Top             =   4800
      Width           =   6375
   End
   Begin VB.CommandButton cmbArh 
      Caption         =   "Архівувати"
      Height          =   387
      Left            =   240
      TabIndex        =   13
      ToolTipText     =   "Архівація супровідного і файла ел.пенсіної справи"
      Top             =   5400
      Width           =   2025
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      DragMode        =   1  'Automatic
      Height          =   345
      Left            =   0
      TabIndex        =   27
      Top             =   7140
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
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   3240
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
         MaxLength       =   12
         TabIndex        =   11
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
         TabIndex        =   10
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
         MaxLength       =   6
         TabIndex        =   9
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label19 
         Caption         =   "Файл"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label18 
         Caption         =   "ПІБ пенсіонера"
         Height          =   315
         Left            =   2760
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "ОР №"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdExportWord 
      Caption         =   "Супровідний лист"
      Height          =   387
      Left            =   240
      TabIndex        =   12
      ToolTipText     =   "Сформувати супровідний лист"
      Top             =   4800
      Width           =   2025
   End
   Begin VB.Frame Frame1 
      Caption         =   "Область та район вибуття"
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   2
      Left            =   240
      TabIndex        =   20
      Top             =   1800
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
         TabIndex        =   22
         Top             =   750
         Width           =   570
      End
      Begin VB.Label Label13 
         Caption         =   "Область"
         Height          =   225
         Left            =   480
         TabIndex        =   21
         Top             =   375
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Область та район прибуття"
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   3
      Left            =   3360
      TabIndex        =   17
      Top             =   1800
      Width           =   5415
      Begin VB.TextBox txtRajTo 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   4320
         MaxLength       =   2
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtOblTo 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   4320
         MaxLength       =   2
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmbSelRajTo 
         Caption         =   "..."
         Height          =   300
         Left            =   4800
         TabIndex        =   8
         ToolTipText     =   "Оберіть район"
         Top             =   720
         Width           =   400
      End
      Begin VB.CommandButton cmbSelOblTo 
         Caption         =   "..."
         Height          =   300
         Left            =   4800
         TabIndex        =   6
         ToolTipText     =   "Оберіть область"
         Top             =   360
         Width           =   400
      End
      Begin VB.Label lblNameObl 
         Alignment       =   2  'Center
         Caption         =   "Оберіть область"
         Height          =   300
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblNameRaj 
         Alignment       =   2  'Center
         Caption         =   "Оберіть район"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   3945
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Дані листа-запиту на висилку"
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   2895
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         Top             =   840
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dataF 
         Height          =   300
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   16646145
         CurrentDate     =   40333
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "№"
         Height          =   195
         Left            =   720
         TabIndex        =   16
         Top             =   840
         Width           =   195
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Запит від "
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   405
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
      Height          =   840
      Left            =   5160
      TabIndex        =   2
      Top             =   600
      Width           =   1770
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4920
      Top             =   240
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
Public myFrmObl As New Form2_
Public myFrmRaj As New Form3

Public NameArhive As String      ' назва архіва 1ls + doc
Public NameDoc As String         ' назва супровідної
Public PathKatalogSend As String ' каталог передачі
Public fso As New FileSystemObject

   Dim myADO As ADODB.Connection
   Dim myRS As ADODB.Recordset
   
Dim hConnection As Long, _
          hOpen As Long, _
          hFile As Long, _
         dwType As Long, _
        dwSeman As Long

'Public appWord As Word.Application, docWord As Word.Document

Private Sub cmbArh_Click()
Dim myCmd As String, tmpNameFile, Quest
      
tmpNameFile = Split(txtFile.Text, ".")
NameArhive = txtOblFrom & txtRajFrom & "_to_" & _
                txtOblTo & txtRajTo & "_" & Right(tmpNameFile(0), 4) & ".rar"

'перевірка наявності вигрузки 1LS
If Not fso.FileExists(PathKatalogSend & txtFile) Then
    MsgBox "Файл вигрузки відсутній", vbCritical, "Архівація не можлива!"
'перевірка наявності супровідної
ElseIf Not fso.FileExists(PathKatalogSend & Left(txtFile, 8) & ".doc") Then
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
    If fso.FileExists(PathKatalogSend & NameArhive) Then _
txtLog.Text = txtLog.Text & "=======" & vbCrLf & Time & " = Архівація пройшла успішно!" & vbCrLf & _
            Time & " = Створено архів " & PathKatalogSend & NameArhive & vbCrLf & "=======" & vbCrLf
End If

End Sub

Private Sub cmbImport_Click()
Dim f As Integer, s As String, fio As String, fio1 As String, FL As String
Dim v As Variant
Dim ss As String, sLS As String
Dim s1 As String, s2 As String
Dim i As Integer, j As Integer
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
ss = ""
sLS = ""

txtLog.Text = Time & " = Відкрито файл " & FL & vbCrLf

Do Until EOF(f)
Line Input #f, s
If Left(s, 3) = "$S," Then ss = s
If Left(s, 3) = "$LS" Then sLS = s
Loop
Close #f

If ss > "" Then
For Each v In Split(ss, ",")
Select Case Left(v, 2)
Case "B="
    txtRajFrom = Mid(v, 5, 2)
Case "F="
    fio = Mid(v, 4, Len(v) - 4)
    fio1 = Dos2Win(fio)
    For i = 1 To 2
    Select Case i
      Case 1: s1 = Chr(161)
              s2 = "і"
      Case 2: s1 = Chr(176)
              s2 = "ї"
    End Select
    j = InStr(1, fio1, s1)
    While j <> 0
      fio1 = Left(fio1, j - 1) + s2 + Mid(fio1, j + 1)
      j = InStr(j + 1, fio1, s1)
    Wend
  Next i
txtPib = StrConv(fio1, vbProperCase)
End Select

Next
End If

    If sLS > "" Then
    txtOsob = Val(Mid(sLS, 5))
End If
txtFile = Dir(FL)

'кнопка Супровідна - активна
'cmdExportWord.Enabled = True
'кнопка Архівувати - активна
'cmbArh.Enabled = True
'текст в лог
txtLog.Text = txtLog.Text & Time & " = Дані успішно імпортовані." & vbCrLf

otvet = MsgBox("Виїзд в межах області??", vbYesNo + vbInformation, "УВАГА")
Select Case otvet
Case vbYes
    txtOblFrom.Text = ReadINI("Viezd", "KodObl", PathFileIni)
    txtOblTo.Text = txtOblFrom.Text
    'cmbSelRajTo.SetFocus
    lblNameObl.Caption = "Хмельницька обл."
    Set myFrmRaj.myFrmVp = Me
    myFrmRaj.Show
Case vbNo
    txtOblFrom = ReadINI("Viezd", "KodOblUa", PathFileIni)
    txtOblTo = ""
    'cmbSelOblTo.SetFocus
        Set myFrmObl.myFrmVp = Me
    myFrmObl.Show
End Select

ErrorHandler:
If Err.Number = 32755 Then
   Exit Sub
End If

End Sub

Private Sub cmbSelOblTo_Click()
    Set myFrmObl.myFrmVp = Me
    myFrmObl.Show
End Sub

Private Sub cmbSelRajTo_Click()
    Set myFrmRaj.myFrmVp = Me
    myFrmRaj.Show
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
    sTmpRajOrObl = txtRajTo & "/"
End If
' якщо код області по УКРАЇНІ
If txtOblFrom = ReadINI("Viezd", "KodOblUa", PathFileIni) Then
    sTmpFtp = Ftp_Kiev
    sTmpKat = Kat_Kiev
    sTmpLog = Login_Kiev
    stmpPass = Pass_Kiev
    sTmpRajOrObl = txtOblTo & "/"
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
   Sleep 500
' встановлюємо цільовий каталог -----------------------------------------
    If (FtpSetCurrentDirectory(hConnection, sTmpKat & sTmpRajOrObl) = False) Then
        ErrorOut Err.LastDllError, "FtpSetCurrentDirectory"
        Exit Sub
    Else
        txtLog.Text = txtLog.Text & Time & " = Каталог для копіювання " & sTmpKat & sTmpRajOrObl & vbCrLf
    End If
'==========================================================================
Sleep 500
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

    If fso.FolderExists(PathKatalogSend & Folder_God_Mes) = False Then
    fso.CreateFolder (PathKatalogSend & Folder_God_Mes)
    fso.CreateFolder (PathKatalogSend & Folder_God_Mes & "\" & "Arhiv")
    Else
        If fso.FolderExists(PathKatalogSend & Folder_God_Mes & "\" & "Arhiv") = False Then
        fso.CreateFolder (PathKatalogSend & Folder_God_Mes & "\" & "Arhiv")
        End If
    End If
'==========================================
    txtLog.Text = txtLog.Text & "=======" & vbCrLf
    'перевірка на наявність вигрузки 1лс
    If fso.FileExists(PathKatalogSend & txtFile.Text) = False Then
    txtLog.Text = txtLog.Text & Time & " @ Відсутній файл вигрузки 1LS " & PathKatalogSend & txtFile.Text & vbCrLf
    Else
    'перенесення вигрузки 1лс в каталог РІК_МІСЯЦЬ
    fso.CopyFile PathKatalogSend & txtFile.Text, PathKatalogSend & Folder_God_Mes & "\" & txtFile.Text, True
    fso.DeleteFile PathKatalogSend & txtFile.Text
    txtLog.Text = txtLog.Text & Time & " = " & txtFile.Text & " перенесено в " & PathKatalogSend & Folder_God_Mes & "\" & vbCrLf
    End If
    '======================================
    Sleep 500
    'перевірка на наявність  супровідної
    If fso.FileExists(PathKatalogSend & NameDoc) = False Then
    txtLog.Text = txtLog.Text & Time & " @ Відсутня супровідна " & PathKatalogSend & NameDoc & vbCrLf
    Else
    'перенесення супровідної в каталог РІК_МІСЯЦЬ
    fso.CopyFile PathKatalogSend & NameDoc, PathKatalogSend & Folder_God_Mes & "\" & NameDoc, True
    fso.DeleteFile PathKatalogSend & NameDoc
    txtLog.Text = txtLog.Text & Time & " = " & NameDoc & " перенесено в " & PathKatalogSend & Folder_God_Mes & "\" & vbCrLf
    End If
    '======================================
    Sleep 500
    'перевірка на наявність АРХІВА
    If fso.FileExists(PathKatalogSend & NameArhive) = False Then
    txtLog.Text = txtLog.Text & Time & " @ Відсутній архів " & PathKatalogSend & NameArhive & vbCrLf
    Else
    'перенесення АРХІВАв каталог РІК_МІСЯЦЬ\arhiv
    fso.CopyFile PathKatalogSend & NameArhive, PathKatalogSend & Folder_God_Mes & "\" & "Arhiv" & "\" & NameArhive
    fso.DeleteFile PathKatalogSend & NameArhive
    txtLog.Text = txtLog.Text & Time & " = " & NameArhive & " перенесено в " & PathKatalogSend & Folder_God_Mes & "\" & "Arhiv" & vbCrLf
    End If
    
Set myADO = New ADODB.Connection
myADO.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & DataBasePath
   
Set myRS = New ADODB.Recordset
myRS.CursorLocation = adUseClient
tmpSql = "INSERT INTO DEP (dep_ins, dep_or, dep_pib, dep_from, dep_to, dep_nfl) " & _
        "VALUES (" & _
        ReadINI("viezd", "ID", PathFileIni) & "," & _
        txtOsob.Text & "," & _
        "'" & txtPib.Text & "'" & "," & _
        "'" & txtRajFrom.Text & "'" & "," & _
        "'" & txtOblTo.Text & txtRajTo.Text & "'" & "," & _
        "'" & txtFile.Text & "'" & ")"

myRS.Open tmpSql, myADO, adOpenDynamic
txtLog.Text = txtLog.Text & Time & " = Запис додано в базу" & vbCrLf

Set myRS = Nothing
Set myRS = New ADODB.Recordset
myRS.CursorLocation = adUseClient
tmpSQL1 = "SELECT * FROM DEP"
myRS.Open tmpSQL1, myADO, adOpenDynamic
txtLog.Text = txtLog.Text & Time & " = Кількість записів в базі " & myRS.RecordCount

myRS.close
Set myRS = Nothing
Set myADO = Nothing

End Sub

Private Sub cmdExportWord_Click()
Dim appWord As Word.Application, _
    docWord As Word.Document, _
    rngCurrent As Word.Range, _
    objTable As Word.Table, _
    s As String, _
    iTBL_Rows As Integer
   
On Error GoTo Err_AccessToWord
    Set appWord = New Word.Application
    
    'Створюєм документ
        Set docWord = appWord.Documents.Add()
        With docWord.PageSetup
            .TopMargin = CentimetersToPoints(1.5)
            .LeftMargin = CentimetersToPoints(2.5)
            .BottomMargin = CentimetersToPoints(1.5)
        End With
        appWord.Visible = False
    '--------------------------------------------
    'колонтітул нижній - дата
With docWord.Sections(1)
    .Footers(wdHeaderFooterPrimary).Range.Text = Date
    .Footers(wdHeaderFooterPrimary).PageNumbers.Add
    .Footers(wdHeaderFooterPrimary).Range.Select
    appWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    appWord.Selection.Font.Name = "Times New Roman"
    appWord.Selection.Font.Size = 8
End With
    
With docWord.ActiveWindow
    .ActivePane.close
    .View = wdPrintView
End With
   
   
    'пишем протокол з БУФЕРА ОБМІНУ
    Set rngCurrent = docWord.Range
    With rngCurrent
        .Collapse Direction:=wdCollapseEnd
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Text = Clipboard.GetText
        .Select
        .Font.Name = "Lucida Console"
        .Font.Size = 8
        .Font.Bold = False
        .InsertParagraphAfter
        .InsertParagraphAfter
        .InsertParagraphAfter
        .InsertParagraphAfter
        .InsertParagraphAfter
        .InsertParagraphAfter
    End With
        
Set rngCurrent = docWord.Range
With rngCurrent
    .InsertParagraphAfter
    .Collapse Direction:=wdCollapseEnd
End With
   '===================================
    
   
    'пишем ЕЛЕКТРОННА ПОШТА
    Set rngCurrent = docWord.Range
    With rngCurrent
        .Collapse Direction:=wdCollapseEnd
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Text = "ЕЛЕКТРОННА ПОШТА"
        .Select
        .Font.Name = "Times New Roman"
        .Font.Size = 16
        .Font.Bold = True
        .InsertParagraphAfter
    End With
        
Set rngCurrent = docWord.Range
With rngCurrent
    .InsertParagraphAfter
    .Collapse Direction:=wdCollapseEnd
End With
        
'додаєм табличку
iTBL_Rows = 1
Set objTable = docWord.Tables.Add(Range:=rngCurrent, NumRows:=iTBL_Rows, NumColumns:=2)
        
With objTable
    .Borders.Enable = False 'не видима таблиця
    .Rows.Height = 10
    .Columns.Width = 250
    .Cell(1, 1).Range.Text = "№ " & Text1 & " від " & dataF    'номер супровідної
    ' шапочка "кому"
    Select Case txtOblFrom.Text
        Case "68"
        .Cell(1, 2).Range.Text = "Управління пенсійного фонду в Хмельницькій області"
        Case "22"
        .Cell(1, 2).Range.Text = "Головне управління Пенсійного фонду України в " & lblNameObl.Caption
    End Select
End With
    
    'перший стовбчик по правому краю, други - по лівому
    With objTable
        .Columns(2).Select
        appWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        .Columns(1).Select
        appWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    End With
        
    Set objTable = Nothing
    '-----------------------------
    'під шапочкою короткий зміст листа
    Set rngCurrent = docWord.Range
    With rngCurrent
        .Collapse Direction:=wdCollapseEnd
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Text = "Про передачу електронної пенсійної справи" & vbCrLf
        .Select
        .Font.Name = "Times New Roman"
        .Font.Size = 12
    End With
    '----------------------------------------
    'тект листа
    Set rngCurrent = docWord.Range
    With rngCurrent
        .InsertParagraphAfter
        .Collapse Direction:=wdCollapseEnd
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Text = "Головне управління Пенсійного фонду України в Хмельницькій області " _
                    & "передає електронну пенсійну справу одержувача пенсії у зв'язку " _
                    & "зі зміною постійного місця проживання:" & vbCrLf
        .Select
        .Font.Size = 14
        .ParagraphFormat.FirstLineIndent = CentimetersToPoints(1.25)
        .ParagraphFormat.Alignment = wdAlignParagraphJustify
    End With
    '----------------------------------------
    'табличка 2 на 5
    Set rngCurrent = docWord.Range
    With rngCurrent
        .InsertParagraphAfter
        .Collapse Direction:=wdCollapseEnd
    End With
        
    Set objTable = docWord.Tables.Add(Range:=rngCurrent, NumRows:=iTBL_Rows + 1, NumColumns:=5)
    With objTable
        .Borders.Enable = True
        .Rows.Height = 10
        .Columns.Width = 60
        'називаєм колонки
        .Cell(1, 1).Range.Text = "№ п/п"
        .Cell(1, 2).Range.Text = "ПІБ"
        .Cell(1, 3).Range.Text = "Область, район вибуття"
        .Cell(1, 4).Range.Text = "Район прибуття"
        .Cell(1, 5).Range.Text = "Назва файлу"
        .Cell(2, 1).Range.Text = "1"
        .Cell(2, 2).Range.Text = txtPib & vbCrLf & "ОР № " & txtOsob 'ПІБ
        .Cell(2, 3).Range.Text = "Хмельницька обл.," & vbCrLf & Select_Raj_Hm(txtRajFrom.Text) & _
                                "  (" & txtOblFrom.Text & txtRajFrom.Text & ")"
        .Cell(2, 4).Range.Text = lblNameRaj & "  (" & txtOblTo & txtRajTo & ")"
        .Cell(2, 5).Range.Text = txtFile
    End With

        objTable.Select
        appWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        appWord.Selection.Font.Size = 14
        appWord.Selection.MoveDown
       
        With objTable
        .AllowAutoFit = True
            .Columns(1).Width = 30
            .Columns(2).Width = 100
            .Columns(3).Width = 140
            .Columns(4).Width = 140
            .Columns(5).Width = 100
            .Columns.Select
            appWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Rows(1).Select
            appWord.Selection.Font.Bold = True
            appWord.Selection.Font.Size = 13
            .Rows(2).Select
            appWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        End With
        
        Set rngCurrent = docWord.Range
        With rngCurrent
            .InsertParagraphAfter
            .Collapse Direction:=wdCollapseEnd
        End With
                
        Set objTable = docWord.Tables.Add(Range:=rngCurrent, NumRows:=1, NumColumns:=2)
        With objTable
            .Borders.Enable = False
            .Rows.Height = 10
            .Columns.Width = 250
            .Cell(1, 1).Range.Text = ReadINI("viezd", "pos", PathFileIni) ' ПІДПИС ПОСАДА
            .Cell(1, 2).Range.Text = ReadINI("viezd", "ln", PathFileIni) ' ПІДПИС ПРІЗВИЩЕ
            .Columns(2).Select
            appWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
            .Cell(1, 2).VerticalAlignment = wdCellAlignVerticalBottom
            .Columns(1).Width = 350
            .Columns(2).Width = 150
        End With
        
            objTable.Select
            appWord.Selection.Font.Size = 14
            appWord.Selection.MoveDown
        
        Set rngCurrent = docWord.Range
        
        With rngCurrent
            .InsertParagraphAfter
            .Collapse Direction:=wdCollapseEnd
        End With

        Set rngCurrent = docWord.Range
        
        With rngCurrent
            .InsertParagraphBefore
            .Collapse Direction:=wdCollapseEnd
            .ParagraphFormat.Alignment = wdAlignParagraphJustify
            .Text = ReadINI("viezd", "prf", PathFileIni) ' ПІДПИС ВИКОНАВЕЦЬ
            .Select
            .Font.Name = "Times New Roman"
            .Font.Size = 13
         End With

        Set objTable = Nothing
        
        NameDoc = Left(txtFile.Text, 8) & ".doc"
        appWord.Visible = True
        docWord.SaveAs PathKatalogSend & NameDoc 'ЗБЕРІГАЄ ФАЙЛ
    
    Set appWord = Nothing
    Set docWord = Nothing
    Set rngCurrent = Nothing
    
    txtLog.Text = txtLog.Text & Time & " = Створено супровідну " & PathKatalogSend & NameDoc & vbCrLf
L_Exit:

    Exit Sub

Err_AccessToWord:
'    AppActivate "Microsoft Access"
    Beep
    MsgBox "The Following Automation Error has occurred:" _
        & vbCrLf & Err.Description & "  " & Err.Number, vbCritical, "Automation Error!"
    Exit Sub
cmdExportWord.Enabled = False
End Sub


Private Sub Form_Load()

Me.Width = 9200
Me.Height = 8000

dataF.Value = Date
cmdExportWord.Enabled = False
cmbArh.Enabled = False

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

End Sub

Private Sub Form_Unload(Cancel As Integer)
InternetCloseHandle hOpen

myFrmMain.viezd.Enabled = True
Set myRS = Nothing
Set myADO = Nothing
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
TextPressEnter (KeyAscii)
End Sub

Private Sub txtFile_KeyPress(KeyAscii As Integer)
TextPressEnter (KeyAscii)
End Sub

Private Sub txtFile_Validate(Cancel As Boolean)
If Len(txtFile.Text) > 0 Then
    cmdExportWord.Enabled = True
    cmbArh.Enabled = True
    ElseIf Len(txtFile.Text) = 0 Then cmdExportWord.Enabled = False And _
                                             cmbArh.Enabled = False
End If
End Sub

Private Sub txtLog_Change()
txtLog.SelStart = Len(txtLog)
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

Public Function TextPressEnter(Key As Integer)
If Key = vbKeyReturn Then
    SendKeys "{Tab}"
End If
End Function

