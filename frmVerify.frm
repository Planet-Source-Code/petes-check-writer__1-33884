VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVerify 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Verify Check data"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   Icon            =   "frmVerify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      Picture         =   "frmVerify.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   75
      Width           =   480
   End
   Begin VB.CommandButton cmdNoGood 
      Caption         =   "Go Back"
      Height          =   510
      Left            =   3045
      Picture         =   "frmVerify.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2610
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Print"
      Default         =   -1  'True
      Height          =   510
      Left            =   2205
      Picture         =   "frmVerify.frx":09CE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2610
      Width           =   795
   End
   Begin MSComctlLib.ListView lvwVerify 
      Height          =   1995
      Left            =   15
      TabIndex        =   0
      Top             =   600
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   3519
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Field"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Data to be printed"
         Object.Width           =   7832
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000003&
      Caption         =   "Verify check data before printing on check.... at this time place the check into the check holder and into the printer."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   540
      Left            =   660
      TabIndex        =   4
      Top             =   45
      Width           =   5640
   End
End
Attribute VB_Name = "frmVerify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Verified As Boolean

Private Sub cmdNoGood_Click()
    Verified = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Verified = True
    Unload Me
End Sub

