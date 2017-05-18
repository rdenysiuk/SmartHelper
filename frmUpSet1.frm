VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUpSet 
   Caption         =   "Журнал оновлень АСОПД"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUpSet1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4785
   ScaleWidth      =   7920
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   5535
      Begin VB.CommandButton cmb_Select 
         Caption         =   "Формувати"
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   360
         Width           =   1380
      End
      Begin MSMask.MaskEdBox txt_from 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##.##.####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt_to 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##.##.####"
         PromptChar      =   "_"
      End
   End
End
Attribute VB_Name = "frmUpSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public myFrmMain As frmMain

Private Sub cmb_Select_Click()
Dim SelUpd
SelUpd = "SELECT "
End Sub

Private Sub Form_Load()
myFrmMain.upd_set.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
myFrmMain.upd_set.Enabled = True
End Sub
