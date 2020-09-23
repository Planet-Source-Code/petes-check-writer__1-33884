VERSION 5.00
Begin VB.Form frmWriteWizard 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Write Check Wizard"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "frmWriteWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmWriteWizard.frx":058A
   ScaleHeight     =   3750
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picWiz 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2670
      Index           =   0
      Left            =   15
      ScaleHeight     =   2640
      ScaleWidth      =   6060
      TabIndex        =   0
      Top             =   1020
      Width           =   6090
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next >"
         Height          =   300
         Index           =   0
         Left            =   2640
         TabIndex        =   1
         Top             =   1935
         Width           =   945
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "WELCOME.   If you write many checks  . . . CheckWriter will make the whole process a little bit easier."
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
      Index           =   4
      Left            =   15
      ScaleHeight     =   2640
      ScaleWidth      =   6060
      TabIndex        =   24
      Top             =   1020
      Width           =   6090
      Begin VB.CommandButton cmdNext 
         Caption         =   "Finish"
         Height          =   300
         Index           =   4
         Left            =   2910
         TabIndex        =   26
         Top             =   2295
         Width           =   945
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "< Back"
         Height          =   300
         Index           =   3
         Left            =   1920
         TabIndex        =   25
         Top             =   2295
         Width           =   945
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "CONGRATULATIONS... you just printed out your check data on your standard wallet check."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Index           =   4
         Left            =   45
         TabIndex        =   27
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
      Index           =   3
      Left            =   15
      ScaleHeight     =   2640
      ScaleWidth      =   6060
      TabIndex        =   15
      Top             =   1020
      Width           =   6090
      Begin VB.CommandButton cmdPrintCheckData 
         Caption         =   "Print Check Data"
         Height          =   795
         Left            =   2145
         Picture         =   "frmWriteWizard.frx":3252
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   570
         Width           =   1515
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "< Back"
         Height          =   300
         Index           =   2
         Left            =   1920
         TabIndex        =   17
         Top             =   2295
         Width           =   945
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next >"
         Height          =   300
         Index           =   3
         Left            =   2910
         TabIndex        =   16
         Top             =   2295
         Width           =   945
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWriteWizard.frx":355C
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Index           =   3
         Left            =   45
         TabIndex        =   19
         Top             =   30
         Width           =   5985
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWriteWizard.frx":35FB
         Height          =   795
         Left            =   60
         TabIndex        =   18
         Top             =   1470
         Width           =   5985
      End
   End
   Begin VB.PictureBox picWiz 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2670
      Index           =   2
      Left            =   15
      ScaleHeight     =   2640
      ScaleWidth      =   6060
      TabIndex        =   10
      Top             =   1020
      Width           =   6090
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   2880
         TabIndex        =   23
         Top             =   1920
         Width           =   3105
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   2880
         TabIndex        =   22
         Top             =   480
         Width           =   3105
      End
      Begin VB.FileListBox File1 
         Height          =   1650
         Left            =   60
         Pattern         =   "*.ckc"
         TabIndex        =   21
         Top             =   480
         Width           =   2805
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next >"
         Height          =   300
         Index           =   2
         Left            =   2910
         TabIndex        =   12
         Top             =   2295
         Width           =   945
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "< Back"
         Height          =   300
         Index           =   1
         Left            =   1920
         TabIndex        =   11
         Top             =   2295
         Width           =   945
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWriteWizard.frx":3717
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Index           =   2
         Left            =   45
         TabIndex        =   13
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
      Left            =   15
      ScaleHeight     =   2640
      ScaleWidth      =   6060
      TabIndex        =   3
      Top             =   1020
      Width           =   6090
      Begin VB.CommandButton cmdBack 
         Caption         =   "< Back"
         Height          =   300
         Index           =   0
         Left            =   1920
         TabIndex        =   5
         Top             =   2295
         Width           =   945
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next >"
         Height          =   300
         Index           =   1
         Left            =   2910
         TabIndex        =   4
         Top             =   2295
         Width           =   945
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWriteWizard.frx":37AE
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   210
         TabIndex        =   9
         Top             =   1335
         Width           =   5745
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "...  Account No: - a place to enter in account numbers.  By default the account number prints above the name/address."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   210
         TabIndex        =   8
         Top             =   840
         Width           =   5745
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "There are 2 fields that are not standard fields on a check:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   45
         TabIndex        =   7
         Top             =   585
         Width           =   5925
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "1. Enter the data in the fields that you would normally hand write.   When all fields are completed click Next."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
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
      Left            =   1185
      TabIndex        =   14
      Top             =   570
      Width           =   4890
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmWriteWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iIndex As Integer
Private Const wizCount As Integer = 4

Private Sub cmdBack_Click(Index As Integer)
iIndex = iIndex - 1
    If iIndex <= 0 Then
        iIndex = 0
        'cmdBack(Index).Enabled = False
    End If
    
    picWiz(iIndex).ZOrder (0)
End Sub

Private Sub cmdNext_Click(Index As Integer)
 Dim i As Long
    With frmCheckWrite
    Select Case Index
        Case 0
            'init
            Load frmCheckWrite
            .InitForm
            .Show
            .SetFocus
            .mnuCheckNew_Click
            Me.Left = frmMain.Width - Me.Width
            frmNavigator.Left = Me.Left
            
        Case 1 'select design
            Drive1.Drive = Mid$(App.path, 1, 3)
            Dir1.path = App.path
            If Len(p_clsOptions.DefaultCheckPath) = 0 Then
                File1.path = Dir1.path
            Else
                File1.path = GetFileName(p_clsOptions.DefaultCheckPath, r_DriverLetterPath)
                'File1.FileName = GetFileName(p_clsOptions.DefaultCheckPath, r_FileNameExt)
                
                For i = 0 To File1.ListCount - 1
                    If UCase(File1.List(i)) = UCase(GetFileName(p_clsOptions.DefaultCheckPath, r_FileNameExt)) Then
                        File1.Selected(i) = True
                        Exit For
                    End If
                Next i
            End If
        Case 2 'print desing sample
            'move to next
        Case 3 'start editing
            'last screen
            Unload frmCheckWrite
        Case 4 'done editing
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


Private Sub cmdPrintCheckData_Click()
    frmCheckWrite.clsPrint.OpenCheckFile = File1.FileName
    frmCheckWrite.mnuCheckPrintCheck_Click
End Sub

Private Sub Dir1_Change()
    File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
    Dir1.path = Drive1.Drive
End Sub

Private Sub Form_Load()
    LblDesc.Caption = "Enter and print data on a stanadard check that you would normally hand write."
    iIndex = 0
End Sub

