VERSION 5.00
Begin VB.Form frmDesignWizard 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Design Wizard"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmDesignWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDesignWizard.frx":058A
   ScaleHeight     =   3750
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picWiz 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2670
      Index           =   0
      Left            =   30
      ScaleHeight     =   2640
      ScaleWidth      =   6060
      TabIndex        =   1
      Top             =   1035
      Width           =   6090
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next >"
         Height          =   300
         Index           =   0
         Left            =   2640
         TabIndex        =   3
         Top             =   1935
         Width           =   945
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmDesignWizard.frx":3291
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Index           =   0
         Left            =   45
         TabIndex        =   2
         Top             =   30
         Width           =   5985
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picWiz 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2670
      Index           =   2
      Left            =   30
      ScaleHeight     =   2640
      ScaleWidth      =   6060
      TabIndex        =   9
      Top             =   1035
      Width           =   6090
      Begin VB.CommandButton cmdPrintDesignSample 
         Caption         =   "Print Design Sample"
         Height          =   675
         Left            =   2100
         Picture         =   "frmDesignWizard.frx":33C6
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Print check design sample"
         Top             =   1125
         Width           =   1575
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next >"
         Height          =   300
         Index           =   2
         Left            =   2910
         TabIndex        =   11
         Top             =   2295
         Width           =   945
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "< Back"
         Height          =   300
         Index           =   1
         Left            =   1920
         TabIndex        =   10
         Top             =   2295
         Width           =   945
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmDesignWizard.frx":3950
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Index           =   2
         Left            =   45
         TabIndex        =   12
         Top             =   30
         Width           =   5985
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picWiz 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2670
      Index           =   5
      Left            =   30
      ScaleHeight     =   2640
      ScaleWidth      =   6060
      TabIndex        =   26
      Top             =   1035
      Width           =   6090
      Begin VB.CheckBox chkSetDefault 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmDesignWizard.frx":3A13
         Height          =   1170
         Left            =   210
         TabIndex        =   30
         Top             =   630
         Width           =   5730
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Finish"
         Height          =   300
         Index           =   5
         Left            =   2910
         TabIndex        =   28
         Top             =   2295
         Width           =   945
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "< Back"
         Height          =   300
         Index           =   4
         Left            =   1920
         TabIndex        =   27
         Top             =   2295
         Width           =   945
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "  CONGRATULATIONS!  You have created a new check design."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   5
         Left            =   45
         TabIndex        =   29
         Top             =   30
         Width           =   5985
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picWiz 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2670
      Index           =   4
      Left            =   30
      ScaleHeight     =   2640
      ScaleWidth      =   6060
      TabIndex        =   18
      Top             =   1035
      Width           =   6090
      Begin VB.CommandButton cmdBack 
         Caption         =   "< Back"
         Height          =   300
         Index           =   3
         Left            =   1920
         TabIndex        =   20
         Top             =   2295
         Width           =   945
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next >"
         Height          =   300
         Index           =   4
         Left            =   2910
         TabIndex        =   19
         Top             =   2295
         Width           =   945
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "c. Enter the coordinate value (in inches) in the Field Coordinates section."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   555
         TabIndex        =   24
         Top             =   1365
         Width           =   5370
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "b. Press the up, down, left, right arrow buttons on your keyboard."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   555
         TabIndex        =   23
         Top             =   1035
         Width           =   5205
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "a. Click and drag the field to the desired coordinate."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   555
         TabIndex        =   22
         Top             =   735
         Width           =   5205
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmDesignWizard.frx":3B89
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   4
         Left            =   45
         TabIndex        =   21
         Top             =   30
         Width           =   5985
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Click Next when you are done modifying your design..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   75
         TabIndex        =   25
         Top             =   1875
         Width           =   5370
      End
   End
   Begin VB.PictureBox picWiz 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2670
      Index           =   3
      Left            =   30
      ScaleHeight     =   2640
      ScaleWidth      =   6060
      TabIndex        =   14
      Top             =   1035
      Width           =   6090
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next >"
         Height          =   300
         Index           =   3
         Left            =   2910
         TabIndex        =   16
         Top             =   2295
         Width           =   945
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "< Back"
         Height          =   300
         Index           =   2
         Left            =   1920
         TabIndex        =   15
         Top             =   2295
         Width           =   945
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmDesignWizard.frx":3C58
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   3
         Left            =   45
         TabIndex        =   17
         Top             =   30
         Width           =   5985
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picWiz 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2670
      Index           =   1
      Left            =   30
      ScaleHeight     =   2640
      ScaleWidth      =   6060
      TabIndex        =   4
      Top             =   1035
      Width           =   6090
      Begin VB.TextBox txtFileName 
         Height          =   315
         Left            =   90
         TabIndex        =   8
         Top             =   585
         Width           =   5805
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "< Back"
         Height          =   300
         Index           =   0
         Left            =   1920
         TabIndex        =   7
         Top             =   2295
         Width           =   945
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next >"
         Height          =   300
         Index           =   1
         Left            =   2910
         TabIndex        =   5
         Top             =   2295
         Width           =   945
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "1. Enter a file name for your check design.  This file name will be the file you choose when prompted for a design."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   45
         TabIndex        =   6
         Top             =   30
         Width           =   5985
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label LblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Desc..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   735
      TabIndex        =   0
      Top             =   600
      Width           =   5370
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDesignWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iIndex As Integer
Private Const wizCount As Integer = 5



Private Sub cmdBack_Click(Index As Integer)
iIndex = iIndex - 1
    If iIndex <= 0 Then
        iIndex = 0
        'cmdBack(Index).Enabled = False
    End If
    
    picWiz(iIndex).ZOrder (0)
End Sub

Private Sub cmdNext_Click(Index As Integer)
    
    With frmCheckRuler
    Select Case Index
        Case 0
            Load frmCheckRuler
            frmCheckRuler.InitForm
            frmCheckRuler.Show
            txtFileName.SetFocus
        Case 1 'set filename
            .mnuDesignNew_Click
            If FileExists(txtFileName & ".ckc") Then
                If MsgBox("Check design file: " & txtFileName & " already exists!" & vbCrLf & "Do you want to overwrite this file?", vbQuestion + vbYesNo, "File Exists") = vbNo Then
                    Exit Sub
                End If
            End If
            .clsHandler.OpenFile = txtFileName.Text
            .mnuDesignSave_Click
        Case 2 'print desing sample
            'move to next
        Case 3 'start editing
            Me.Left = frmMain.Width - Me.Width
            frmNavigator.Left = Me.Left
            frmCheckRuler.SetFocus
            .mnuDesignDesign_Click
        Case 4 'done editing
            .mnuDesignSave_Click
        Case 5 'set default
            If chkSetDefault.Value = vbChecked Then
                p_clsOptions.DefaultCheckPath = .clsHandler.OpenFile
                p_clsOptions.SaveOptions
            End If
            picWiz(0).ZOrder (0)
            Unload Me
            Exit Sub
    End Select
    End With
    iIndex = iIndex + 1
    If iIndex >= wizCount Then
        iIndex = wizCount
        'cmdNext(Index).Enabled = False
    End If
    
    picWiz(iIndex).ZOrder (0)
    
End Sub

Private Sub cmdPrintDesignSample_Click()
    frmCheckRuler.mnuDesignPrintDesignSample_Click
End Sub

Private Sub Form_Load()

    LblDesc.Caption = "A check design lets you specify the coordinates of the fields that will print on your standard pre-printed check."
    iIndex = 0
    
    
End Sub

