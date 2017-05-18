VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmZNDI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Статистика по НДІ"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   Icon            =   "frmZNDI1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   6360
   Begin VB.CheckBox Check1 
      Caption         =   "Без врахування інспектора"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton cmd_Form 
      Caption         =   "Формувати"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   4440
      Width           =   2535
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   3015
      Left            =   480
      TabIndex        =   6
      Top             =   1200
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5318
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   12582912
      ForeColorFixed  =   -2147483634
      AllowBigSelection=   0   'False
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).BandIndent=   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.ComboBox cmbYear1 
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Text            =   "2010"
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox cmbMonth1 
      Height          =   315
      ItemData        =   "frmZNDI1.frx":000C
      Left            =   480
      List            =   "frmZNDI1.frx":000E
      TabIndex        =   2
      Text            =   "1"
      Top             =   720
      Width           =   975
   End
   Begin VB.ComboBox cmbYear2 
      Height          =   315
      Left            =   4800
      TabIndex        =   1
      Text            =   "2010"
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox cmbMonth2 
      Height          =   315
      Left            =   3600
      TabIndex        =   0
      Text            =   "1"
      Top             =   720
      Width           =   975
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
      Left            =   3480
      TabIndex        =   5
      Top             =   360
      Width           =   2415
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
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "frmZNDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public myFrmMain As frmMain


Private Sub Check1_Click()
With Grid1

If Check1.Value = 0 Then
    .Clear
    .Rows = 2
    .Cols = 3
    .FormatString = "Звітна дата | Кількість | Інспектор "
    .ColWidth(0) = 1200
    .ColWidth(1) = 1000
    .ColWidth(2) = 1000
    
Else
    .Clear
    .Rows = 2
    .Cols = 2
    .FormatString = "Звітна дата | Кількість"
    .ColWidth(0) = 1200
    .ColWidth(1) = 1000
End If

End With
End Sub

Private Sub cmd_Form_Click()
Dim SelNdiKol, kolgroup As String

If Len(cmbMonth1) = 1 Then cmbMonth1 = "0" & cmbMonth1
If Len(cmbMonth2) = 1 Then cmbMonth2 = "0" & cmbMonth2

SelNdiKol = "SELECT convert(varchar(6), ndi_date, 112) as [Звітна дата], SUM(ndi_kol) As [Кількість] " & _
  "FROM [SSPZ].[dbo].[ndi] WHERE convert (varchar(6), ndi_date, 112) BETWEEN " & cmbYear1 & cmbMonth1 & " AND " & cmbYear2 & cmbMonth2 & _
  "GROUP BY convert(varchar(6), ndi_date, 112) "
  
kolgroup = "SELECT convert(varchar(6), ndi_date, 112) as [Звітна дата], SUM(ndi_kol) As [Кількість], ndi_prog as [Інспектор] " & _
  "FROM [SSPZ].[dbo].[ndi] WHERE convert (varchar(6), ndi_date, 112) BETWEEN " & cmbYear1 & cmbMonth1 & " AND " & cmbYear2 & cmbMonth2 & _
  " GROUP BY convert(varchar(6), ndi_date, 112), ndi_prog ORDER BY convert(varchar(6), ndi_date, 112)"
  
  
ConnectToDataBase
If Check1.Value = 1 Then
    myRS.Open SelNdiKol, myADO, adOpenStatic
Else
    myRS.Open kolgroup, myADO, adOpenStatic
'MsgBox kolgroup
End If
If myRS.RecordCount >= 1 Then
    Set Grid1.Recordset = myRS
Else
    MsgBox "За вказаний період дані відсутні.", vbInformation, "Пхе"
End If

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

'With Grid1
'    .Rows = 2
'    .Cols = 2
'    .FormatString = "Звітна дата | Кількість"
'    .ColWidth(0) = 1200
'    .ColWidth(1) = 1000
'End With
Check1.Value = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
myFrmMain.zndi.Enabled = True
End Sub

