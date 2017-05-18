VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2_ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Довідник областей"
   ClientHeight    =   8325
   ClientLeft      =   1095
   ClientTop       =   435
   ClientWidth     =   4455
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleMode       =   0  'User
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmObl.frx":0000
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   13150
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   16
      RowDividerStyle =   6
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "obl_kod"
         Caption         =   "Код"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "obl_name"
         Caption         =   "Найменування"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3000,189
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   7920
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   7680
      Width           =   4335
   End
End
Attribute VB_Name = "Form2_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public myFrmVp As frmVP
Public myFrmRaj As Form3

Private Sub DataGrid1_DblClick()
Dim mR As String, sC As String, myTmp, myTmp1
With DataGrid1
    mR = .Row
    .Col = 0
    myTmp = .Text
    mR = .Row
    .Col = 1
    myTmp1 = .Text
End With

With myFrmVp
    .txtOblTo = myTmp
    .lblNameObl = myTmp1
End With

'Set myFrmRaj.myFrmObl = Me
'myFrmRaj.Show
Unload Me
'Form3.Show
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
DataGrid1.AllowUpdate = True
With Label1
    .Caption = "УВАГА!!! РЕДАГУВАННЯ"
    .FontBold = True
    .ForeColor = &HC0&
    Label2.FontBold = True
End With
End If

If KeyAscii = 27 Then
DataGrid1.AllowUpdate = False
With Label1
    .Caption = "Редагування - ENTER"
    .FontBold = False
    .ForeColor = &HC00000
    Label2.FontBold = False
End With
End If
End Sub


Private Sub Form_Load()
    
    Dim tmpSql As String
   ConnectToDataBase
   tmpSql = "SELECT obl_kod, obl_name FROM obl ORDER BY obl_kod"
   myRS.Open tmpSql, myADO, adOpenDynamic

Set DataGrid1.DataSource = myRS
frmMain.Enabled = False
Label1.Caption = "Редагування - ENTER"
Label2.Caption = "Завершити редагування - ESC"
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
Set myADO = Nothing
Set myRS = Nothing
Unload Me
End Sub

