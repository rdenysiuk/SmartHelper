VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Довідник районів"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5595
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   5595
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   13150
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Appearance      =   0
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   16
      RowDividerStyle =   5
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
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "raj_name"
         Caption         =   "Назва"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "dd.MM.yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "raj_kod"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   4004,788
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   705,26
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   8040
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   7680
      Width           =   5175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public myFrmVp As frmVP, tmpObl
Public myFrmObl As Form2_

Private Sub DataGrid1_DblClick()
Dim mR As String, mC As String, myTmp, myTmp1
With DataGrid1
    'mR = .Row
    .Col = 0
    myTmp = .Text
    .Col = 1
    myTmp1 = .Text
'    myFrmVp.txtRajTo = myTmp1
'    myFrmVp.lblNameRaj = myTmp
End With

With myFrmVp
    .txtRajTo = myTmp1
    .lblNameRaj = myTmp
    .cmdExportWord.Enabled = True
'    .cmdExportWord.SetFocus
    .cmbArh.Enabled = True
End With

Unload Me
myFrmVp.cmdExportWord.SetFocus
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
DataGrid1.AllowUpdate = True
With Label1
    .Caption = "УВАГА!!! РЕДАГУВАННЯ" '& vbCrLf & "ESC - завершити редагування"
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
   
   Dim tmpSQL As String
   tmpObl = myFrmVp.txtOblTo.Text
   
ConnectToDataBase
   tmpSQL = "SELECT raj_name, raj_kod FROM raj WHERE raj_obl='" & tmpObl & "' ORDER BY raj_name"
   myRS.Open tmpSQL, myADO, adOpenDynamic
   
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


