VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Лінивчик"
   ClientHeight    =   7020
   ClientLeft      =   6165
   ClientTop       =   3195
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00008000&
   Icon            =   "lazy1.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7000
   ScaleMode       =   0  'User
   ScaleWidth      =   6400
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   6660
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   11748
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Надіслати звіти"
      TabPicture(0)   =   "lazy1.frx":0ECA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Виїзд пенсіонера"
      TabPicture(1)   =   "lazy1.frx":0EE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Про програму"
      TabPicture(2)   =   "lazy1.frx":0F02
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Архівація REP-файлів"
      TabPicture(3)   =   "lazy1.frx":0F1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

