VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmZVP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Звіти по вигрузках 1LS"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11340
   Icon            =   "frmZVP1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7048.1
   ScaleMode       =   0  'User
   ScaleWidth      =   11865.47
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   11880
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Звіт по датах"
      TabPicture(0)   =   "frmZVP1.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdSelect"
      Tab(0).Control(1)=   "Text1"
      Tab(0).Control(2)=   "cmbMonth2"
      Tab(0).Control(3)=   "cmbYear2"
      Tab(0).Control(4)=   "cmbMonth1"
      Tab(0).Control(5)=   "cmbYear1"
      Tab(0).Control(6)=   "chkUser"
      Tab(0).Control(7)=   "cmdExToExcel"
      Tab(0).Control(8)=   "ListView1"
      Tab(0).Control(9)=   "Label1"
      Tab(0).Control(10)=   "Label2"
      Tab(0).Control(11)=   "Label3"
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Пошук виїзду"
      TabPicture(1)   =   "frmZVP1.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblTarget"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdSearch"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtTarget"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "pib"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "osob"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "GridViewSearch"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridViewSearch 
         Height          =   4935
         Left            =   240
         TabIndex        =   18
         Top             =   1560
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   8705
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorFixed  =   16711680
         ForeColorFixed  =   -2147483628
         AllowBigSelection=   0   'False
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "Номер | ПІБ пенсіонера | Прог | Дата | Звідки | Куди(код) | Куди(назва) "
         BandDisplay     =   1
         _NumberOfBands  =   1
         _Band(0).BandIndent=   1
         _Band(0).Cols   =   7
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.OptionButton osob 
         Caption         =   "За номером ОР"
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
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton pib 
         Caption         =   "За ПІБ пенсіонера"
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
         Height          =   495
         Left            =   2400
         TabIndex        =   15
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtTarget 
         Height          =   315
         Left            =   2520
         TabIndex        =   14
         Top             =   1080
         Width           =   5055
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Пошук"
         Height          =   375
         Left            =   7800
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Формувати"
         Height          =   375
         Left            =   -74520
         TabIndex        =   8
         Top             =   5880
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   -71520
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   5400
         Width           =   1335
      End
      Begin VB.ComboBox cmbMonth2 
         Height          =   315
         Left            =   -71400
         TabIndex        =   6
         Text            =   "1"
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox cmbYear2 
         Height          =   315
         Left            =   -70200
         TabIndex        =   5
         Text            =   "2010"
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cmbMonth1 
         Height          =   315
         ItemData        =   "frmZVP1.frx":0044
         Left            =   -74520
         List            =   "frmZVP1.frx":0046
         TabIndex        =   4
         Text            =   "1"
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox cmbYear1 
         Height          =   315
         Left            =   -73320
         TabIndex        =   3
         Text            =   "2010"
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox chkUser 
         Caption         =   "Без врахування інспектора"
         Height          =   255
         Left            =   -74520
         TabIndex        =   2
         Top             =   6360
         Width           =   2415
      End
      Begin VB.CommandButton cmdExToExcel 
         Caption         =   "Експорт в Excel"
         Height          =   375
         Left            =   -71160
         TabIndex        =   1
         Top             =   5880
         Width           =   2655
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3975
         Left            =   -74520
         TabIndex        =   9
         Top             =   1320
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   7011
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
         NumItems        =   4
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
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   1687
         EndProperty
      End
      Begin VB.Label lblTarget 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Початкова дата (місяць, рік)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74640
         TabIndex        =   12
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Кінцева дата (місяць, рік)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -71520
         TabIndex        =   11
         Top             =   480
         Width           =   2415
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
         Left            =   -73320
         TabIndex        =   10
         Top             =   5400
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmZVP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public myFrmMain As frmMain
Private Declare Function ActivateKeyboardLayout& Lib "user32" (ByVal HKL As Long, _
ByVal Flags As Long)

Private Sub cmbMonth1_Click()
If cmbMonth1 = 12 Then
    cmbMonth2 = 1
    cmbYear2 = cmbYear1 + 1
End If

End Sub

Private Sub cmdExToExcel_Click()
Dim myXL As Excel.Application, _
    myWB As Excel.Workbook, _
    myWS As Excel.Worksheet
    
Dim I As Long ' Для перебора строк в ListView
Dim y As Integer ' Для перебора колонок в ListView. У меня их 3
Dim z As Long 'Для перебора строк в Excel. Поскольку начинаю не с первой, а с третьей.

Set myXL = New Excel.Application
Set myWB = myXL.Workbooks.Open(App.Path & "\Report\" & "RepCount1ls.xlt")
Set myWS = myWB.Worksheets(1)
z = 8
myXL.Visible = True

For I = 1 To Me.ListView1.ListItems.Count
    myWS.Cells(z, 2) = ListView1.ListItems(I).Index
    myWS.Cells(z, 3) = ListView1.ListItems(I)
    For y = 1 To 3
        'myWS.Cells(1, y + 1) = ListView1.ColumnHeaders(y)
        myWS.Cells(z, y + 3) = ListView1.ListItems(I).SubItems(y)
    Next y
z = z + 1
Next I
'==========================================
'With myWS
'For I = 0 To myRS.Fields.Count - 1
'    .Range("A1").Offset(0, I).Value = myRS.Fields(I).Name
'    Next
'    .Range("A2").CopyFromRecordset (myRS)
'End With
'==========================================
'For I = 0 To myRS.Fields.Count - 1
'    myWS.Cells(1, I + 1).Value = myRS.Fields(I).Name
'Next
'myWS.Range(myWS.Cells(1, 1), _
'    myWS.Cells(1, myRS.Fields.Count)).Font.Bold = True
'myWS.Range("A2").CopyFromRecordset myRS, 100, 100

'MsgBox myRS.RecordCount
Set myXL = Nothing
Set myWB = Nothing
Set myWS = Nothing

End Sub

Private Sub cmdSearch_Click()
Dim SS As String, S1 As String


SS = "SELECT dep_or as [Номер], dep_pib as [ПІБ пенсіонера], dep_ins as [Прог], dep_date as [Дата], dep_from as [Звідки],dep_to as [Куди(код)], raj_name as [Куди(назва)] FROM DEP LEFT JOIN raj on (left(dep_to,2)=raj_obl) and (right(dep_to,2)=raj_kod) WHERE "
If osob = True Then
    S1 = "dep_or Like '" & txtTarget.Text & "%' ORDER by dep_or"
Else
    S1 = "dep_pib Like '" & txtTarget.Text & "%' ORDER by dep_pib"
End If

ConnectToDataBase

myRS.Open SS & S1, myADO, adOpenStatic
If myRS.RecordCount = 0 Then
    MsgBox "Нічого не знайдено.", vbInformation, " Atention!!! "
Else
    Set GridViewSearch.Recordset = myRS
End If
End Sub

Private Sub cmdSelect_Click()

If Len(cmbMonth1) <> 2 Then cmbMonth1 = "0" & cmbMonth1
If Len(cmbMonth2) <> 2 Then cmbMonth2 = "0" & cmbMonth2
CreateCount1ls
End Sub



Private Sub Form_Load()
Dim iMonth As Integer, iYear As Integer

For iMonth = 1 To 12
    cmbMonth1.AddItem iMonth
    cmbMonth2.AddItem iMonth
Next
    
For iYear = 2010 To Year(Date) + 1
    cmbYear1.AddItem iYear
    cmbYear2.AddItem iYear
Next

cmbMonth1 = Month(Date)
If cmbMonth1 = 12 Then
    cmbMonth2 = 1
    cmbYear2 = Year(Date) + 1
Else
    cmbMonth2 = cmbMonth1 + 1
    cmbYear2 = Year(Date)
End If

cmbYear1 = Year(Date)
chkUser.Value = 1

lblTarget = "Ведіть № ОР"
txtTarget.MaxLength = 6
osob.Value = True

With GridViewSearch
    .ColWidth(0) = 650
    .ColWidth(1) = 2700
    .ColWidth(2) = 450
    .ColWidth(3) = 1350
    .ColWidth(4) = 600
    .ColWidth(5) = 810
    .ColWidth(6) = 4200

End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set myRS = Nothing
Set myADO = Nothing
Unload Me
End Sub

Private Sub osob_Click()
'ListView2.ListItems.Clear
lblTarget = "Ведіть № ОР"
With txtTarget
    .Text = ""
'    .SetFocus
End With

'GridViewSearch.Clear
End Sub

Private Sub pib_Click()
'ListView2.ListItems.Clear
lblTarget = "Ведіть ПІБ пенсіонера"

With txtTarget
    .Text = ""
    .SetFocus
End With
ActivateKeyboardLayout &H4220422, 3
GridViewSearch.Clear
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

If KeyCode = 13 Then
Call cmdSearch_Click
End If


End Sub

Public Sub CreateCount1ls()
Dim zapros As String, _
      item As ListItem, _
        AA, Beg, TheEnD

ConnectToDataBase

If chkUser.Value = 0 Then
 zapros = "SELECT   convert (varchar(6),dep_date, 112) as Dataa, dep_ins, Count(dep_ins) AS Kol From Dep" & _
 " WHERE convert (varchar(6),dep_date, 112) BETWEEN " & cmbYear1 & cmbMonth1 & " and " & cmbYear2 & cmbMonth2 & _
 " GROUP BY convert (varchar(6),dep_date, 112), dep_ins" & _
 " ORDER BY  convert (varchar(6),dep_date, 112), dep_ins;"
 myRS.Open zapros, myADO, adOpenStatic

 ListView1.ListItems.Clear

 Do While Not myRS.EOF
    Set item = ListView1.ListItems.Add(, , Format("01." & Right(myRS("Dataa"), 2) & "." & Left(myRS("Dataa"), 4), "mmmm yyyy"))
    item.SubItems(1) = (myRS("dep_ins"))
    item.SubItems(2) = (myRS("Kol"))
    AA = AA + myRS("Kol")
    myRS.MoveNext
 Loop
 Text1.Text = AA

Else
 
 zapros = "SELECT   convert (varchar(6),dep_date, 112) as Dataa, Count(dep_ins) AS Kol From Dep" & _
 " WHERE convert (varchar(6),dep_date, 112) BETWEEN " & cmbYear1 & cmbMonth1 & " and " & cmbYear2 & cmbMonth2 & _
 " GROUP BY convert (varchar(6),dep_date, 112)" & _
 " ORDER BY  convert (varchar(6),dep_date, 112);"
 myRS.Open zapros, myADO, adOpenStatic

 ListView1.ListItems.Clear

 Do While Not myRS.EOF
    Set item = ListView1.ListItems.Add(, , Format("01." & Right(myRS("Dataa"), 2) & "." & Left(myRS("Dataa"), 4), "mmmm yyyy"))
    item.SubItems(2) = (myRS("Kol"))
    AA = AA + myRS("Kol")
    myRS.MoveNext
 Loop
 Text1.Text = AA
End If
End Sub
