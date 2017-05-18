VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAddType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Довідник техніки"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9780
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddType1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9780
   Begin VB.TextBox txt_search 
      Height          =   350
      Left            =   240
      TabIndex        =   18
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CheckBox chk_filtr 
      Caption         =   "  Застосувати умови фільтру"
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Caption         =   " Фільтр "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   9375
      Begin VB.CommandButton cmb_filtr 
         Caption         =   "Застосувати"
         Height          =   375
         Left            =   7080
         TabIndex        =   14
         Top             =   280
         Width           =   2055
      End
      Begin VB.ComboBox cmb_type_f 
         Height          =   330
         Left            =   4560
         TabIndex        =   13
         Top             =   280
         Width           =   2175
      End
      Begin VB.TextBox txt_inv_f 
         Height          =   350
         Left            =   1680
         TabIndex        =   12
         Top             =   280
         Width           =   1935
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Інвентраний №"
         Height          =   350
         Left            =   240
         TabIndex        =   17
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Тип"
         Height          =   350
         Left            =   4080
         TabIndex        =   16
         Top             =   300
         Width           =   330
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Додати пристрій "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1455
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   9375
      Begin VB.CommandButton cmbAddDev 
         Caption         =   "Додати"
         Height          =   495
         Left            =   7080
         TabIndex        =   4
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txt_kab 
         Height          =   350
         Left            =   4560
         TabIndex        =   3
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txt_model 
         Height          =   350
         Left            =   1680
         TabIndex        =   2
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox cmb_type 
         Height          =   330
         Left            =   4560
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txt_inv 
         Height          =   350
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Кабінет"
         Height          =   210
         Left            =   3840
         TabIndex        =   11
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Модель"
         Height          =   210
         Left            =   960
         TabIndex        =   10
         Top             =   840
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Тип"
         Height          =   210
         Left            =   4080
         TabIndex        =   9
         Top             =   360
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Інвентраний №"
         Height          =   210
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid AllDev 
      Height          =   3375
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   12582912
      ForeColorFixed  =   -2147483634
      GridColorFixed  =   16777215
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).BandIndent=   1
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frmAddType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public myFrmMain As frmMain

Dim sSort As String



Private Sub AllDev_Click()
With AllDev
    If .MouseRow = 0 Then
     .Col = .MouseCol
     .Sort = flexSortStringAscending
    'Else
    ' .CellFontBold = True
    End If
  End With
End Sub

Private Sub AllDev_KeyPress(KeyAscii As Integer)

With txt_search
    .Visible = True
    .SetFocus
End With

End Sub

Private Sub chk_filtr_Click()
If chk_filtr.Value = 1 Then
    Filtr (True)
Else
    Filtr (False)
End If
End Sub

Private Sub cmb_filtr_Click()
Dim Dev, Filtr As String

Dev = "SELECT di_nom as Інвентарний, dev_name as Тип, di_model as Модель, di_kab as Кабінет " & _
    "FROM dev_inv INNER JOIN device on di_dev=dev_id " 'WHERE di_nom=" & txt_inv_f.Text '& _
    '" and di_dev = " & cmb_type_f.ItemData(cmb_type_f.ListIndex)

If Len(txt_inv_f.Text) = 0 And Len(cmb_type_f.Text) > 0 Then _
    Filtr = "WHERE di_dev = " & Val(cmb_type_f.ItemData(cmb_type_f.ListIndex))

If Len(txt_inv_f.Text) > 0 And Len(cmb_type_f.Text) = 0 Then _
    Filtr = "WHERE di_nom = " & txt_inv_f.Text
    
If Len(txt_inv_f.Text) > 0 And Len(cmb_type_f.Text) > 0 Then _
    Filtr = "WHERE di_nom = " & txt_inv_f.Text & " and di_dev = " & Val(cmb_type_f.ItemData(cmb_type_f.ListIndex))
    
ConnectToDataBase
myRS.Open Dev & Filtr, myADO, adOpenStatic
Set AllDev.Recordset = myRS
'MsgBox Filtr
End Sub

Private Sub cmbAddDev_Click()
Dim addDev As String
On Error GoTo erradd

addDev = "INSERT INTO dev_inv (di_nom, di_dev, di_model, di_kab) VALUES " & _
    "(" & txt_inv.Text & "," & cmb_type.ItemData(cmb_type.ListIndex) & ",'" & _
    txt_model.Text & "'," & txt_kab.Text & ");"


If Len(txt_inv.Text) <> 0 And Len(cmb_type.Text) <> 0 And Len(txt_model.Text) <> 0 And Len(txt_kab.Text) <> 0 Then
    ConnectToDataBase
    myRS.Open addDev, myADO, adOpenStatic
    Call SpAllDev
    txt_inv.Text = ""
    txt_model.Text = ""
    txt_kab.Text = ""
Else
    MsgBox "Одне або декілька полів не заповнено.", vbExclamation + vbOKOnly, "Atention"
End If

erradd:
    If Err.Number = -2147217873 Then
        MsgBox "Пристрій з інвентарнийм номером " & txt_inv.Text & " вже заведений", vbExclamation, "Atention"
    ElseIf Err.Number <> 0 Then
        MsgBox Err.Number & vbTab & Err.Description, vbExclamation, "Warning"
    End If

End Sub


Private Sub Form_Load()
Dim sp As String

Filtr (False)
sp = "SELECT * FROM device ORDER BY dev_name"

Call SpAllDev

ConnectToDataBase
myRS.Open sp, myADO, adOpenStatic

With cmb_type
Do While Not myRS.EOF
    .AddItem myRS("dev_name").Value
    cmb_type_f.AddItem myRS("dev_name").Value
    .ItemData(.NewIndex) = myRS("dev_id").Value
    cmb_type_f.ItemData(cmb_type_f.NewIndex) = myRS("dev_id").Value
    myRS.MoveNext
Loop
End With

cmb_type.ListIndex = 2
txt_search.Visible = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
myFrmMain.dov.Enabled = True
End Sub

Private Sub SpAllDev()
Dim I As Integer
Dim selAllDev As String
On Error GoTo spAlldeverr

With AllDev
.Clear
    .ColWidth(0) = 1500
    .ColWidth(1) = 2000
    .ColWidth(2) = 2000
    .ColWidth(3) = 900
End With

selAllDev = "SELECT di_nom as Інвентарний, dev_name as Тип, di_model as Модель, di_kab as Кабінет FROM dev_inv INNER JOIN device on di_dev=dev_id"
ConnectToDataBase
myRS.Open selAllDev, myADO, adOpenStatic

Set AllDev.DataSource = myRS

Exit Sub

spAlldeverr:
MsgBox Err.Description
Resume Next

End Sub


Private Sub Filtr(mode As Boolean)
If mode = False Then
    cmb_filtr.Enabled = False
    txt_inv_f.Enabled = False
    txt_inv_f.Text = ""
    cmb_type_f.Enabled = False
    Frame2.Enabled = False
    Call SpAllDev
Else
    cmb_filtr.Enabled = True
    txt_inv_f.Enabled = True
    cmb_type_f.Enabled = True
    Frame2.Enabled = True
End If

End Sub

Private Sub txt_search_Change()
Dim I, j As Integer

If Len(txt_search.Text) > 4 Then
    For I = 0 To AllDev.Rows - 1
        If InStr(AllDev.TextMatrix(I, 0), txt_search.Text) Then
        AllDev.CellFontBold = True
        AllDev.Row = I
        End If
    Next I
Else
    AllDev.CellFontBold = False
End If

End Sub

Private Sub txt_search_LostFocus()
With txt_search
    .Text = ""
    .Visible = False
End With
End Sub
