VERSION 5.00
Begin VB.Form frmSelectPrintType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select to Print"
   ClientHeight    =   705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "frmSelectPrintType.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   660
      Left            =   3720
      Picture         =   "frmSelectPrintType.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   15
      Width           =   1200
   End
   Begin VB.CommandButton cmdDesign 
      Caption         =   "Check Design"
      Height          =   660
      Left            =   2475
      Picture         =   "frmSelectPrintType.frx":0894
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print Check Design Sample"
      Top             =   15
      Width           =   1200
   End
   Begin VB.CommandButton cmdData 
      Caption         =   "Check Data"
      Height          =   660
      Left            =   2475
      Picture         =   "frmSelectPrintType.frx":09DE
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Check Data"
      Top             =   15
      Width           =   1200
   End
   Begin VB.CommandButton cmdSample 
      Caption         =   "Check Sample"
      Height          =   660
      Left            =   1245
      Picture         =   "frmSelectPrintType.frx":0F68
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print Check Sample"
      Top             =   15
      Width           =   1200
   End
   Begin VB.CommandButton cmdHolder 
      Caption         =   "Check Holder"
      Height          =   660
      Left            =   15
      Picture         =   "frmSelectPrintType.frx":14F2
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Check Holder"
      Top             =   15
      Width           =   1200
   End
End
Attribute VB_Name = "frmSelectPrintType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum enmPrintChecks
    chk_None = 0
    chk_Holder = 1
    chk_Sample = 2
    chk_Data = 3
    chk_Design = 4
End Enum

Public Response As enmPrintChecks


Private Sub cmdCheckHolder_Click()

End Sub
Public Sub InitForm(Optional pbPrintData As Boolean = False, Optional pbPrintDesign As Boolean = False)
    
    Me.Response = chk_None
    cmdData.Visible = pbPrintData
    cmdDesign.Visible = pbPrintDesign
    
    
End Sub

Private Sub cmdData_Click()
    Me.Response = chk_Data
    Unload Me
End Sub

Private Sub cmdDesign_Click()
    Me.Response = chk_Design
    Unload Me
End Sub

Private Sub cmdExit_Click()

    Me.Response = chk_None
    Unload Me
End Sub

Private Sub cmdHolder_Click()
    Me.Response = chk_Holder
    Unload Me
End Sub

Private Sub cmdSample_Click()
    Me.Response = chk_Sample
    Unload Me
End Sub

