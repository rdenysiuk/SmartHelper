VERSION 5.00
Object = "{41AFDA5A-831B-4895-865A-7FB6994EB548}#6.0#0"; "rsp-zip-compress150.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Надіслати звіти"
   ClientHeight    =   7020
   ClientLeft      =   6765
   ClientTop       =   4110
   ClientWidth     =   6345
   Icon            =   "frmSR.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7000
   ScaleMode       =   0  'User
   ScaleWidth      =   6500
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   6120
      Width           =   6135
   End
   Begin VB.CommandButton cmdArhiv 
      Caption         =   "Надіслати"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Надіслати звіти"
      Top             =   5715
      Width           =   2400
   End
   Begin VB.ComboBox cmbNameArh 
      Height          =   315
      Left            =   360
      TabIndex        =   7
      ToolTipText     =   "Вкажіть назву архіва. Можна редагувати"
      Top             =   5760
      Width           =   2520
   End
   Begin VB.Frame Frame1 
      Caption         =   " Оберіть район для відправки"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5535
      Begin VB.ComboBox cmbRaj 
         Height          =   315
         ItemData        =   "frmSR.frx":08CA
         Left            =   1440
         List            =   "frmSR.frx":08CC
         Sorted          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Код району"
         Top             =   480
         Width           =   2175
      End
   End
   Begin RSPZipCompress150.RSPZip RSPZip1 
      Left            =   5400
      Top             =   1800
      _ExtentX        =   979
      _ExtentY        =   979
   End
   Begin MSComctlLib.ListView List1 
      Height          =   2175
      Left            =   360
      TabIndex        =   6
      Top             =   2520
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3836
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
         Text            =   "Назва"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Розмір"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Дата"
         Object.Width           =   3175
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   6765
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Назва архіву"
      Height          =   195
      Left            =   480
      TabIndex        =   11
      Top             =   5520
      Width           =   1005
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Оберіть район."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   360
      TabIndex        =   10
      ToolTipText     =   "Строка стану програми"
      Top             =   4920
      Width           =   5535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Каталог для пошуку звітів - "
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Поштовий каталог -"
      Height          =   195
      Left            =   840
      TabIndex        =   4
      Top             =   2160
      Width           =   1530
   End
   Begin VB.Label lblKatalog 
      AutoSize        =   -1  'True
      Caption         =   "Не визначено"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2760
      TabIndex        =   3
      Top             =   1920
      Width           =   2145
   End
   Begin VB.Label LblKatalogS 
      AutoSize        =   -1  'True
      Caption         =   "Не визначено"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2760
      TabIndex        =   2
      Top             =   2160
      Width           =   2145
   End
End
Attribute VB_Name = "frmSR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public myFrmMain As frmMain
Dim fso As New FileSystemObject, _
    fld As Folder, _
     MI As Integer, _
PriznZV, _
  NomerFl As Integer, _
NameArhSend As String

Const sRepNAR As String = "RepNar", _
      sRepMAS As String = "RepMas", _
         sF2A As String = "F2a", _
        sFssN As String = "Fssn"

Private Sub cmbNameArh_Click()
Dim tmpSql
PriznZV = Split(cmbNameArh, "_")
PriznZV = Left(PriznZV(0), (InStr(1, PriznZV(0), "6")) - 1)
ConnectToDataBase
tmpSql = "SELECT * FROM Send WHERE Sr_Raj = " & cmbRaj & _
         " and Sr_NameZv = '" & PriznZV & _
         "' and Sr_Month = " & MI & _
         " and Sr_Year = " & Right(Year(Date), 2)

myRS.Open tmpSql, myADO, adOpenDynamic
    NomerFl = myRS.RecordCount
Text1.Text = Text1.Text & NomerFl
End Sub

Private Sub cmbRaj_Click()
Dim kol As Integer, Ves As Currency
List1.ListItems.Clear

Katalogs (True)

Ves = FindFile(Folder_DB_Raj & "base" & cmbRaj & "\post_raj", "*.rpr", kol)
If kol > 0 Then
    Label1.Caption = "В " & Val(Right(cmbRaj, 2)) & "-му районі знайдено " & _
                        kol & " файлів (" & Razmer(Ves) & ")" & " для відправки."
    Label1.ForeColor = &H8000&
    CreateNameArh
    cmdArhiv.Enabled = True
Else
    Label1.Caption = "В " & Val(Right(cmbRaj, 2)) & "-му районі Не знайдено файлів." & _
                        vbCrLf & "Відправка не можлива."
    Label1.ForeColor = &H80&
    
    cmbNameArh.Enabled = False
    cmdArhiv.Enabled = False

End If

ProgressBar1.Value = 0

End Sub

Private Sub cmdArhiv_Click()
Dim comando As String, _
     tmpSql As String

If NomerFl = 0 Then
    NameArhSend = cmbNameArh
Else
    NameArhSend = cmbNameArh & "_" & NomerFl + 1
End If

comando = "<set-zip-temp-path=c:\>" _
        & "<include-system-and-hidden-files>" _
        & "<zip-compression-mode=add-to-zipfile>" _
        & "<compression-level=9>" _
        & "<directory-with-the-files-to-compress=" & Folder_DB_Raj & "BASE" & cmbRaj & "\post_raj\>" _
        & "<destination-directory=" & Folder_DB_Raj & "BASE" & cmbRaj & "\post_raj\>" _
        & "<destination-zipfile=" & NameArhSend & ".zip>" _
        & "<files-selection=*.rpr>"
ConnectToDataBase

tmpSql = "INSERT INTO Send (Sr_Raj, Sr_NameZv, Sr_Month, Sr_Year, Sr_nomFl, Sr_Date, Sr_Ins, Sr_PathFl) " & _
         "Values (" & cmbRaj & ",'" & PriznZV & "'," & MI & "," & Right(Year(Date), 2) & _
         "," & NomerFl + 1 & ",'" & Date & "'" & ",'" & ReadINI("viezd", "ID", PathFileIni) & _
         ",'" & Folder_DB_Raj & "BASE" & cmbRaj & "\post_raj\" & NameArhSend & ".zip' );"
Text1.Text = tmpSql
'RSPZip1.RSPZipCompress (comando)
'myRS.Open tmpSql, myADO, adOpenDynamic

End Sub

Private Sub Form_Load()
Dim tmpOper1, _
    tmpOper2, _
    tmpOper3, _
    tmpOper4, _
      tmpSql, _
     tmpSQL1, _
         All

Me.Width = 6500
Me.Height = 7500


If ReadINI("SendReport", "Oper1", PathFileIni) <> 0 Then _
    tmpOper1 = 1
    All = tmpOper1

If ReadINI("SendReport", "Oper2", PathFileIni) <> 0 Then
    tmpOper2 = 2
    If All <> "" Then
        All = All & " OR oper_nom=" & tmpOper2
    Else
        All = tmpOper2
    End If
End If

If ReadINI("SendReport", "Oper3", PathFileIni) <> 0 Then
    tmpOper3 = 3
    If All <> "" Then
        All = All & " OR oper_nom=" & tmpOper3
    Else
        All = tmpOper3
    End If
End If

If ReadINI("SendReport", "Oper4", PathFileIni) <> 0 Then
    tmpOper4 = 4
    If All <> "" Then
        All = All & " OR oper_nom=" & tmpOper4
    Else
        All = tmpOper4
    End If
End If

tmpSQL1 = "WHERE oper_nom = " & All & ";"
    
cmbNameArh.Enabled = False
cmdArhiv.Enabled = False
cmbRaj.Clear

ConnectToDataBase
tmpSql = "SELECT oper_raj FROM oper " & tmpSQL1
myRS.Open tmpSql, myADO, adOpenDynamic
cmbRaj.Clear

Do While Not myRS.EOF
    cmbRaj.AddItem myRS("oper_raj").Value
    myRS.MoveNext
Loop
myRS.close

End Sub

Private Sub Form_Unload(Cancel As Integer)
myFrmMain.RepSend.Enabled = True
End Sub

Private Function FindFile(ByVal sFol As String, sFile As String, _
                                        nFiles As Integer) As Currency
   Dim tFld As Folder, _
       tFil As File, _
   FileName As String, _
       item As ListItem, _
      File1 As String
               
   On Error GoTo Catch 'end function
   
   Set fld = fso.GetFolder(sFol)
   FileName = Dir(fso.BuildPath(fld.Path, sFile), vbNormal Or vbHidden Or vbSystem Or vbReadOnly)
   While Len(FileName) <> 0
      File1 = fso.BuildPath(fld.Path, FileName)
      FindFile = FindFile + FileLen(File1) '(FSO.BuildPath(fld.Path, FileName))
      nFiles = nFiles + 1
        Set item = List1.ListItems.Add(, , FileName)    ' Load ListBox
            item.SubItems(1) = Razmer(FileLen(File1))
            item.SubItems(2) = FileSystem.FileDateTime(File1)
      FileName = Dir()  ' Get next file
      DoEvents
    'File1 = ""
    Label1.Caption = "Зачекайте. Йде побудова списку файлів."
   Wend
   
Exit Function
Catch:  Set item = List1.ListItems.Add(, , "Немає доступу")    ' Load ListBox
        item.SubItems(1) = "до каталога"
        item.SubItems(2) = sFol
    'Resume Next
End Function


Private Sub List1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
SortListView List1, ColumnHeader
End Sub

Private Sub RSPZip1_Finished(ReturnCode As Long, ReturnDescription As String)
Dim k As Folder, FL As String
Set k = fso.GetFolder(Folder_DB_Raj & "BASE" & cmbRaj & "\POST_RAJ\")

If fso.FileExists(Folder_DB_Raj & "BASE" & cmbRaj & "\POST_RAJ\*.rpr") = False Then
    FL = Dir(fso.BuildPath(k.Path, "*.rpr"), vbNormal)
    While Len(FL) <> 0
        Label1.Caption = "Видалення файлу " & FL
        fso.DeleteFile fso.BuildPath(k.Path, FL), True
        DoEvents
        FL = Dir()
        Sleep 100
    Wend
End If
        
With Label1
    .Caption = "Архівування завершено" 'ReturnCode & ": " & ReturnDescription
    Select Case ReturnCode
        Case 0
        .Caption = .Caption & " УСПІШНО."
        Case 606
        .Caption = .Caption & " невдачею. :-(" & vbCrLf & "Не має доступу до каталогу звітів"
        .ForeColor = &H80&
        Case 12
        .Caption = .Caption & " невдачею. :-(" & vbCrLf & "Не знайдено файлів для архівації"
        .ForeColor = &H80&
    End Select
    .Caption = .Caption & " Зачекайте." & vbCrLf & "Копіювання " & NameArhSend & _
                ".zip => " & Folder_Post & cmbRaj
    'ReturnCode & ": " & ReturnDescription
End With
    
    DoEvents
fso.CopyFile Folder_DB_Raj & "Base" & cmbRaj & "\post_raj\" & NameArhSend & ".zip", _
             Folder_Post & cmbRaj & "\" & NameArhSend & ".zip", True
    DoEvents
Label1.Caption = "Копіювання завершено."

If fso.FileExists(Folder_Post & cmbRaj & "\" & NameArhSend & ".zip") = True Then
    MsgBox "Файл " & Folder_Post & cmbRaj & "\" & NameArhSend & ".zip" & vbCrLf & _
            "поставлено до черги відправлення.", vbInformation, Me.Caption
Else
    MsgBox "Помилка невідомого характеру", vbCritical
End If

List1.ListItems.Clear
ProgressBar1.Value = 0
End Sub

Private Sub RSPZip1_Progress(Progress As Long)
    ProgressBar1.Value = Progress
End Sub

Private Sub RSPZip1_Status(Value As Long)

With Label1
    If Value = 0 Then
        .Caption = "Завершено."
    End If
   
    If (Value = 1) Then
        .Caption = "Побудова списку файлів..."
    End If
    
    If (Value = 2) Then
        .Caption = "Зачекайте йде архівація звітів..."
    End If
End With

End Sub

Public Sub CreateNameArh()
Dim Konec
MI = Month(Date)
    If Day(Date) >= 23 Then MI = MI + 1
    Konec = cmbRaj & "_" & MI & "." & Right(Year(Date), 2)
    With cmbNameArh
        .Enabled = True
        .Clear
        .AddItem sRepNAR & Konec
        .AddItem sRepMAS & Konec
        .AddItem sF2A & Konec
        .AddItem sFssN & Konec
    End With
End Sub
Private Function Katalogs(prizn As Boolean)

If prizn = False Then
    With lblKatalog
        .Caption = "Не визначено"
        .ForeColor = &H80& ' темно-червоний
    End With
    With LblKatalogS
        .Caption = "Не визначено"
        .ForeColor = &H80& ' темно-червоний
    End With

Else
    With lblKatalog
        .Caption = Folder_DB_Raj & "BASE" & cmbRaj & "\POST_RAJ"
        .ForeColor = &H8000& ' зелений
    End With
    With LblKatalogS
        .Caption = Folder_Post & cmbRaj
        .ForeColor = &H8000& ' зелений
    End With
End If

End Function
Public Sub SortListView(ByVal lvw As ListView, _
   ByVal colHdr As ColumnHeader)
  With lvw
  ' установка режима сортировки для указанной колонки
    .SortKey = colHdr.Index - 1
    .Sorted = True
  ' изменение сортировки меняется между
  ' "по возрастанию" и "по уменьшению"
    .SortOrder = 1 Xor .SortOrder
End With
End Sub
Private Function Razmer(a)
    If a < 1000 Then
        Razmer = a
    ElseIf a >= 1000 And a < 1000000 Then
        Razmer = Round(a / 1024, 2) & " Кб"
    ElseIf a >= 1000000 Then
        Razmer = Round(a / 1048576, 2) & " Мб"
    End If
End Function

