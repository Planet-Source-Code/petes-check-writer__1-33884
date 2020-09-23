VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4920
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6150
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   10
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   4455
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4935
      TabIndex        =   2
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3705
      TabIndex        =   1
      Top             =   4455
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   0
      Left            =   165
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   420
      Width           =   5685
      Begin VB.Frame fraSample1 
         Caption         =   "Set defaults values"
         Height          =   3075
         Left            =   30
         TabIndex        =   5
         Top             =   45
         Width           =   5490
         Begin VB.CommandButton cmdGetPrinter 
            Height          =   300
            Left            =   5115
            Picture         =   "frmOptions.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Select printer"
            Top             =   1875
            Width           =   300
         End
         Begin VB.CommandButton cmdGetDesign 
            Height          =   300
            Left            =   5115
            Picture         =   "frmOptions.frx":051E
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Select check design"
            Top             =   630
            Width           =   300
         End
         Begin VB.TextBox txtDefaultPrinter 
            Height          =   285
            Left            =   165
            TabIndex        =   14
            Top             =   1875
            Width           =   4950
         End
         Begin VB.TextBox txtDefaultDesign 
            Height          =   285
            Left            =   165
            TabIndex        =   12
            Top             =   630
            Width           =   4950
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmOptions.frx":0A30
            Height          =   645
            Left            =   165
            TabIndex        =   16
            Top             =   930
            Width           =   5280
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmOptions.frx":0AE6
            Height          =   645
            Left            =   150
            TabIndex        =   17
            Top             =   2220
            Width           =   5280
         End
         Begin VB.Label Label2 
            Caption         =   "Printer:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   150
            TabIndex        =   13
            Top             =   1665
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Check Design:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   150
            TabIndex        =   11
            Top             =   405
            Width           =   2250
         End
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4245
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7488
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Defaults"
            Key             =   "Defaults"
            Object.ToolTipText     =   "Defaults Values"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Design"
            Key             =   "Design"
            Object.ToolTipText     =   "Set Options for designing checks"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Write"
            Key             =   "Write"
            Object.ToolTipText     =   "Set Options for writing checks"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   180
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   465
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Set options for writing checks"
         Height          =   3660
         Left            =   30
         TabIndex        =   9
         Top             =   45
         Width           =   5535
         Begin VB.CheckBox chkEnterTab 
            Caption         =   "Use [ENTER] key to move from field to field"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   180
            TabIndex        =   24
            Top             =   2865
            Width           =   5145
         End
         Begin VB.TextBox txtFullName 
            Height          =   285
            Left            =   195
            TabIndex        =   22
            Top             =   1845
            Width           =   5115
         End
         Begin VB.CheckBox chkAutoComplete 
            Caption         =   "Auto Complete"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   210
            TabIndex        =   15
            Top             =   405
            Width           =   5145
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmOptions.frx":0B85
            Height          =   825
            Left            =   150
            TabIndex        =   25
            Top             =   3180
            Width           =   5280
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmOptions.frx":0C0E
            Height          =   825
            Left            =   180
            TabIndex        =   23
            Top             =   2160
            Width           =   5280
         End
         Begin VB.Label Label10 
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   195
            TabIndex        =   21
            Top             =   1620
            Width           =   2685
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmOptions.frx":0CCF
            Height          =   825
            Left            =   180
            TabIndex        =   18
            Top             =   720
            Width           =   5280
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   180
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   450
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Set design options"
         Height          =   3330
         Left            =   30
         TabIndex        =   26
         Top             =   45
         Width           =   5550
         Begin VB.CheckBox chkEditOnOpen 
            Caption         =   "Edit = ON when New or Open check design"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   165
            TabIndex        =   28
            Top             =   480
            Width           =   4830
         End
         Begin VB.TextBox txtMoveIncrement 
            Height          =   300
            Left            =   210
            TabIndex        =   27
            Text            =   "0.10"
            Top             =   1830
            Width           =   510
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmOptions.frx":0DC1
            Height          =   645
            Left            =   165
            TabIndex        =   32
            Top             =   720
            Width           =   5280
         End
         Begin VB.Label Label7 
            Caption         =   "Move Field Increment"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   180
            TabIndex        =   31
            Top             =   1590
            Width           =   2325
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "In edit mode when using the arrow keys... the value entered above will be the distance the field will move with each keystroke."
            Height          =   645
            Left            =   195
            TabIndex        =   30
            Top             =   2190
            Width           =   5280
         End
         Begin VB.Label Label9 
            Caption         =   "inches"
            Height          =   225
            Left            =   765
            TabIndex        =   29
            Top             =   1920
            Width           =   690
         End
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    p_clsOptions.Apply
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdGetDesign_Click()
    txtDefaultDesign.Text = p_clsOptions.SelectDesign()
End Sub

Private Sub cmdGetPrinter_Click()
    txtDefaultPrinter.Text = p_clsOptions.SelectPrinter()
End Sub

Private Sub cmdOK_Click()
    p_clsOptions.Apply
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
    
   
End Sub
Private Sub Form_Load()
    'center the form
    'Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    p_clsOptions.LoadOptions
End Sub

Private Sub tbsOptions_Click()
    
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Enabled = True
            picOptions(i).ZOrder (0)
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
    
End Sub

