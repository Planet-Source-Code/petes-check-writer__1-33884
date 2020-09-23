VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmVerifyPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Verify Check data"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   Icon            =   "frmVerifyPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   30
      Picture         =   "frmVerifyPrint.frx":0442
      ScaleHeight     =   1215
      ScaleWidth      =   570
      TabIndex        =   3
      Top             =   45
      Width           =   570
   End
   Begin VB.CommandButton cmdNoGood 
      Caption         =   "Go Back"
      Height          =   510
      Left            =   3045
      Picture         =   "frmVerifyPrint.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3330
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Print"
      Default         =   -1  'True
      Height          =   510
      Left            =   2205
      Picture         =   "frmVerifyPrint.frx":09CE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3330
      Width           =   795
   End
   Begin MSComctlLib.ListView lvwVerify 
      Height          =   1995
      Left            =   15
      TabIndex        =   0
      Top             =   1320
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   3519
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   4
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
      Picture         =   "frmVerifyPrint.frx":0B18
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000003&
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
      Height          =   1215
      Left            =   600
      TabIndex        =   4
      Top             =   45
      Width           =   5640
   End
End
Attribute VB_Name = "frmVerifyPrint"
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

Private Sub Form_Load()
    Label1.Caption = "Verify check data before printing on check. If correct:" & vbCrLf & "1. Place the check into the check holder." & vbCrLf & "2. Place the check holder into the printer." & vbCrLf & "3. Click the Print button or press the [ENTER] key"
    
End Sub
