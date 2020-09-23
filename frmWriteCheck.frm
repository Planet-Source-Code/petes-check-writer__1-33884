VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCheckWrite 
   Caption         =   "Write Check"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8145
   Icon            =   "frmWriteCheck.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3255
   ScaleWidth      =   8145
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      FillColor       =   &H00FFFFC0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   0
      ScaleHeight     =   2970
      ScaleWidth      =   7905
      TabIndex        =   5
      Top             =   15
      Width           =   7935
      Begin VB.CommandButton cmdHistory 
         Height          =   300
         Left            =   5205
         Picture         =   "frmWriteCheck.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1005
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   750
         Left            =   -15
         Picture         =   "frmWriteCheck.frx":0954
         ScaleHeight     =   750
         ScaleWidth      =   4260
         TabIndex        =   15
         Top             =   0
         Width           =   4260
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   4275
         TabIndex        =   4
         Top             =   2505
         Width           =   3465
      End
      Begin VB.TextBox txtMemo 
         Height          =   300
         Left            =   1125
         TabIndex        =   2
         Top             =   1710
         Width           =   3990
      End
      Begin VB.TextBox txtAccountNo 
         Height          =   300
         Left            =   1125
         TabIndex        =   3
         Top             =   2025
         Width           =   4005
      End
      Begin VB.TextBox txtAmount 
         Height          =   300
         Left            =   6270
         TabIndex        =   1
         Top             =   1005
         Width           =   1530
      End
      Begin VB.TextBox txtPayTo 
         Height          =   300
         Left            =   1065
         TabIndex        =   0
         Top             =   1005
         Width           =   4275
      End
      Begin MSComCtl2.DTPicker DTPicker 
         Height          =   315
         Left            =   5280
         TabIndex        =   6
         Top             =   390
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         Format          =   22675457
         CurrentDate     =   36827
      End
      Begin VB.Line Line1 
         X1              =   4290
         X2              =   7725
         Y1              =   2460
         Y2              =   2460
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   240
         Left            =   3750
         TabIndex        =   14
         Top             =   2535
         Width           =   1050
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Memo:"
         Height          =   240
         Left            =   165
         TabIndex        =   13
         Top             =   1740
         Width           =   1050
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Account No:"
         Height          =   240
         Left            =   165
         TabIndex        =   12
         Top             =   2040
         Width           =   1050
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Dollars"
         Height          =   285
         Left            =   6615
         TabIndex        =   11
         Top             =   1410
         Width           =   1095
      End
      Begin VB.Label lblNumText 
         BackStyle       =   0  'Transparent
         Height          =   285
         Left            =   1050
         TabIndex        =   10
         Top             =   1395
         Width           =   5535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount: $"
         Height          =   240
         Left            =   5535
         TabIndex        =   9
         Top             =   1035
         Width           =   870
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Pay To The Order Of:"
         Height          =   480
         Left            =   150
         TabIndex        =   8
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         Height          =   240
         Left            =   4785
         TabIndex        =   7
         Top             =   420
         Width           =   465
      End
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   2910
      Left            =   135
      Top             =   225
      Width           =   7920
   End
   Begin VB.Menu mnuCheck 
      Caption         =   "&Check"
      Begin VB.Menu mnuCheckNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuCheckDesign 
         Caption         =   "&Design"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuCheckOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCheckSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckPrint 
         Caption         =   "&Print"
         Begin VB.Menu mnuCheckPrintCheck 
            Caption         =   "&Check Data"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnuCheckPrintHolder 
            Caption         =   "Check &Holder"
         End
         Begin VB.Menu mnuCheckPrintSample 
            Caption         =   "&Sample Check"
         End
      End
      Begin VB.Menu mnuCheckSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckClose 
         Caption         =   "C&lose"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuViewWizards 
         Caption         =   "&Wizards..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Window"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
      Begin VB.Menu mnuWindowBar1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmCheckWrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public clsPrint As clsWriteCheck


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If p_clsOptions.EnterTab And KeyCode = 13 Then
        SendKeys "{TAB}"
        KeyCode = 0
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, frmMain
End Sub

Private Sub mnuHelpContents_Click()
Call HtmlHelp(0, App.HelpFile, HH_DISPLAY_TOPIC, ByVal 0&)
End Sub

Private Sub mnuViewWizards_Click()
    Load frmNavigator
    frmNavigator.InitForm wiz_Write
    frmNavigator.Show , frmMain
End Sub

Private Sub mnuWindowArrangeIcons_Click()
  frmMain.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowCascade_Click()
  frmMain.Arrange vbCascade
End Sub

Private Sub mnuWindowNewWindow_Click()
Dim NewWindow As New frmCheckWrite
  
  Load NewWindow
  NewWindow.InitForm
  NewWindow.Show
  Set NewWindow = Nothing
  
End Sub

Private Sub mnuWindowTileHorizontal_Click()
  frmMain.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowTileVertical_Click()
  frmMain.Arrange vbTileVertical
End Sub



Public Sub InitForm()


    Set clsPrint = New clsWriteCheck
    Set clsPrint.Form = Me
    clsPrint.Mode = md_None
    DTPicker.Value = Format(Now, "mm/dd/yyyy")
    'txtPayTo.SetFocus
    clsPrint.SetOpenFile p_clsOptions.DefaultCheckPath
    
    Me.Width = 8265
    Me.Height = 3945
    
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Form_Activate()
    ShowToolbar tb_Write
    
    If Not (clsPrint Is Nothing) Then
        SetMDIStatusBarMode clsPrint.Mode
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not (clsPrint Is Nothing) Then
        clsPrint.Mode = md_None
        Set clsPrint = Nothing
    End If
    
    ShowToolbar tb_Main
End Sub



Private Sub mnuCheckClose_Click()
    Unload Me
End Sub

Private Sub mnuCheckDesign_Click()
  Load frmCheckRuler
  frmCheckRuler.InitForm
  frmCheckRuler.Show
End Sub

Public Sub mnuCheckNew_Click()
    clsPrint.NewCheck
    txtPayTo.SetFocus
End Sub

Public Sub mnuCheckPrintCheck_Click()
    clsPrint.PrintData
End Sub

Private Sub mnuCheckPrintHolder_Click()
    PrintHolder
End Sub

Private Sub mnuCheckPrintSample_Click()
    PrintHolder True
End Sub

Private Sub mnuViewOptions_Click()
    ShowOptions
End Sub

Private Sub txtAccountNo_KeyPress(KeyAscii As Integer)
    If p_clsOptions.EnterTab And KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtAmount_Change()
lblNumText.Caption = modNumToText.NumToWord(txtAmount.Text)
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    If p_clsOptions.EnterTab And KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtMemo_KeyPress(KeyAscii As Integer)
    If p_clsOptions.EnterTab And KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If p_clsOptions.EnterTab And KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPayTo_KeyPress(KeyAscii As Integer)
    If p_clsOptions.EnterTab And KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPayTo_LostFocus()
    clsPrint.LoadLogData
End Sub
