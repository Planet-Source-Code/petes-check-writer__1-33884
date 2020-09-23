VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCheckRuler 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000003&
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   1605
   ClientWidth     =   9585
   Icon            =   "frmCheckRuler.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmCheckRuler.frx":0442
   ScaleHeight     =   5685
   ScaleWidth      =   9585
   Begin VB.PictureBox picFieldCor 
      BackColor       =   &H8000000C&
      Height          =   600
      Left            =   705
      ScaleHeight     =   540
      ScaleWidth      =   6630
      TabIndex        =   16
      Top             =   4695
      Visible         =   0   'False
      Width           =   6690
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   5070
         TabIndex        =   2
         Top             =   225
         Width           =   615
      End
      Begin VB.TextBox txtYTop 
         Height          =   285
         Left            =   2880
         TabIndex        =   1
         Top             =   225
         Width           =   615
      End
      Begin VB.TextBox txtXLeft 
         Height          =   285
         Left            =   585
         TabIndex        =   0
         Top             =   225
         Width           =   615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   4545
         TabIndex        =   23
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "inches"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   5730
         TabIndex        =   22
         Top             =   270
         Width           =   540
      End
      Begin VB.Label lblFieldName 
         BackStyle       =   0  'Transparent
         Caption         =   "Field Coordinates"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   90
         TabIndex        =   21
         Top             =   0
         Width           =   3630
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "inches"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   3525
         TabIndex        =   20
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Y Top"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   2340
         TabIndex        =   19
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "inches"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   1245
         TabIndex        =   18
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "X Left"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   60
         TabIndex        =   17
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.PictureBox picHandle 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   0
      Left            =   8880
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   15
      Top             =   2385
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picHandle 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   1
      Left            =   8880
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picHandle 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   2
      Left            =   8910
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   13
      Top             =   2880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picHandle 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   3
      Left            =   8880
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   12
      Top             =   3150
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picHandle 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   4
      Left            =   8895
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   11
      Top             =   3405
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picHandle 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   5
      Left            =   8895
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   10
      Top             =   3675
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picHandle 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   6
      Left            =   8895
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   9
      Top             =   3945
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picHandle 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   7
      Left            =   8895
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   4215
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   5
      Top             =   5385
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6094
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5106
            Picture         =   "frmCheckRuler.frx":8E6B4
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5106
            Picture         =   "frmCheckRuler.frx":8E81C
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox HRuler 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   585
      ScaleHeight     =   495
      ScaleWidth      =   8835
      TabIndex        =   3
      Top             =   -30
      Width           =   8835
   End
   Begin VB.PictureBox VRuler 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   4920
      Left            =   0
      ScaleHeight     =   4920
      ScaleWidth      =   570
      TabIndex        =   4
      Top             =   465
      Width           =   570
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      Picture         =   "frmCheckRuler.frx":8E984
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   -15
      ScaleHeight     =   510
      ScaleWidth      =   315
      TabIndex        =   7
      Top             =   0
      Width           =   315
   End
   Begin VB.Label lblPosition 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Account Info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Index           =   0
      Left            =   1230
      TabIndex        =   24
      Tag             =   "AccountInfo"
      Top             =   825
      Width           =   2880
   End
   Begin VB.Label lblPosition 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Index           =   6
      Left            =   5040
      TabIndex        =   30
      Tag             =   "Name"
      Top             =   3810
      Width           =   3915
   End
   Begin VB.Label lblPosition 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Memo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Index           =   5
      Left            =   1725
      TabIndex        =   29
      Tag             =   "Memo"
      Top             =   3555
      Width           =   2595
   End
   Begin VB.Label lblPosition 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Amount in text and 00/100 ---"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Index           =   4
      Left            =   1185
      TabIndex        =   28
      Tag             =   "AmountText"
      Top             =   2445
      Width           =   6330
   End
   Begin VB.Label lblPosition 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$ Amount ---"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Index           =   3
      Left            =   7485
      TabIndex        =   27
      Tag             =   "AmountValue"
      Top             =   2010
      Width           =   1440
   End
   Begin VB.Label lblPosition 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "** Pay to the order of **"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Index           =   2
      Left            =   1860
      TabIndex        =   26
      Tag             =   "PayTo"
      Top             =   2010
      Width           =   5040
   End
   Begin VB.Label lblPosition 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Index           =   1
      Left            =   6315
      TabIndex        =   25
      Tag             =   "Date"
      Top             =   1395
      Width           =   1440
   End
   Begin VB.Menu mnuCheck 
      Caption         =   "&Design"
      Begin VB.Menu mnuDesignNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuDesignOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuDesignDesign 
         Caption         =   "&Edit"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuDesignSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuDesignSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuCheckSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesignPrint 
         Caption         =   "&Print"
         Begin VB.Menu mnuDesignPrintDesignSample 
            Caption         =   "&Design Sample"
         End
         Begin VB.Menu mnuDesignPrintHolder 
            Caption         =   "&Check Holder"
         End
         Begin VB.Menu mnuDesignPrintSampleCheck 
            Caption         =   "&Sample Check"
         End
      End
      Begin VB.Menu mnuCheckSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesignExit 
         Caption         =   "C&lose"
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
Attribute VB_Name = "frmCheckRuler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Xf1 As Single
Dim Yf1 As Single
Dim bStart As Boolean
Private bKeyMove As Boolean
Public clsHandler As clsFormRuler


Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, frmMain
End Sub

Private Sub mnuHelpContents_Click()
Call HtmlHelp(0, App.HelpFile, HH_DISPLAY_TOPIC, ByVal 0&)
End Sub

Private Sub mnuViewWizards_Click()
    Load frmNavigator
    frmNavigator.InitForm wiz_Design
    frmNavigator.Show , frmMain
End Sub

Private Sub mnuWindowArrangeIcons_Click()
  frmMain.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowCascade_Click()
  frmMain.Arrange vbCascade
End Sub

Private Sub mnuWindowNewWindow_Click()
  Dim NewWindow As New frmCheckRuler
  
  Load NewWindow
  NewWindow.InitForm
  NewWindow.Show
        
    
End Sub

Private Sub mnuWindowTileHorizontal_Click()
  frmMain.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowTileVertical_Click()
  frmMain.Arrange vbTileVertical
End Sub

Public Sub InitForm()
    
    Set clsHandler = New clsFormRuler
    Set clsHandler.Form = Me
    clsHandler.Initialize
    clsHandler.DragInit
    
    If p_clsOptions.EditOnOpen Then
        clsHandler.DesignMode = True
    End If
        
End Sub

Private Sub Form_Activate()
    ShowToolbar tb_Design
    
    If Not (clsHandler Is Nothing) Then
        SetMDIStatusBarMode clsHandler.Mode
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
 Dim bUpdate As Boolean
 
    If bKeyMove Then
        Exit Sub
    End If
    
    If (Not clsHandler.CurrCtl Is Nothing) Then
        clsHandler.CurrCtl.Enabled = False
        Select Case KeyCode
            Case 37 'Left afrow
                clsHandler.CurrCtl.Left = clsHandler.CurrCtl.Left - (p_clsOptions.MoveFieldSize * 1440)
                bUpdate = True
            Case 38 'Up arrow
                clsHandler.CurrCtl.Top = clsHandler.CurrCtl.Top - (p_clsOptions.MoveFieldSize * 1440)
                bUpdate = True
            Case 39 'Right Arrow
                clsHandler.CurrCtl.Left = clsHandler.CurrCtl.Left + (p_clsOptions.MoveFieldSize * 1440)
                bUpdate = True
            Case 40 'Down Arrow
                clsHandler.CurrCtl.Top = clsHandler.CurrCtl.Top + (p_clsOptions.MoveFieldSize * 1440)
                bUpdate = True
        End Select
        If bUpdate = True Then
            clsHandler.CurrCtl.Enabled = True
            clsHandler.DesignChanged = True
            clsHandler.ShowHandles False
        End If
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    clsHandler.Form_MouseDown Button
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    clsHandler.Form_MouseMove x, y
End Sub



Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    clsHandler.Form_MouseUp Button
End Sub

Private Sub Form_Resize()
    If Not (clsHandler Is Nothing) Then
        clsHandler.Initialize True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    
    
    If Not (clsHandler Is Nothing) Then
        clsHandler.Mode = md_None
        
        If mnuDesignDesign.Checked Then
            mnuDesignDesign_Click
        End If
                
        Set clsHandler = Nothing
    End If
    
    ShowToolbar tb_Main
    
End Sub

Private Sub HRuler_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'clsHandler.HandlePositionIndicator 0, -1
    clsHandler.HandlePositionIndicator (HRuler.Left + x), (HRuler.Top + y)
End Sub


Private Sub Label7_Click()

End Sub

Private Sub lblPosition_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
     lblPosition(Index).ToolTipText = "Left = " & TwipToInch(lblPosition(Index).Left - piRulerOffsetX) & " Top = " & TwipToInch(lblPosition(Index).Top - piRulerOffsetY)
    
     
    
    If Button = vbLeftButton And clsHandler.DesignMode Then
        clsHandler.DragBegin lblPosition(Index)
    End If

    clsHandler.DesignChanged = True
End Sub

Private Sub lblPosition_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    clsHandler.x = x
    clsHandler.y = y
      
    
    clsHandler.HandlePositionIndicator (lblPosition(Index).Left + x), (lblPosition(Index).Top + y)
    
    lblPosition(Index).ToolTipText = "Left = " & TwipToInch(lblPosition(Index).Left - piRulerOffsetX) & " Top = " & TwipToInch(lblPosition(Index).Top - piRulerOffsetY)
    sbStatusBar.Panels(2).Text = lblPosition(Index).ToolTipText
    sbStatusBar.Panels(1).Text = lblPosition(Index).Tag
End Sub

Private Sub lblPosition_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    clsHandler.IsMoving = False
End Sub

Public Sub mnuDesignDesign_Click()

Dim i As Integer

    clsHandler.DesignMode = Not mnuDesignDesign.Checked
    mnuDesignDesign.Checked = clsHandler.DesignMode
    
    picFieldCor.Visible = clsHandler.DesignMode
    clsHandler.MDIForm.tbDesign.Buttons("Edit").Value = IIf(clsHandler.DesignMode = True, tbrPressed, tbrUnpressed)
    
End Sub

Private Sub mnuDesignExit_Click()
    
    
    Unload Me
    
End Sub

Public Sub mnuDesignNew_Click()
    clsHandler.NewDesign
    
End Sub

Public Sub mnuDesignOpen_Click()
    clsHandler.OpenDesign
   
End Sub


Public Sub mnuDesignPrintDesignSample_Click()
     clsHandler.PrintDesign
End Sub

Private Sub mnuDesignPrintHolder_Click()
   PrintHolder
End Sub

Private Sub mnuDesignPrintSampleCheck_Click()
    PrintHolder True
End Sub

Public Sub mnuDesignSave_Click()
    clsHandler.SaveDesign
End Sub

Private Sub mnuDesignSaveAs_Click()
    clsHandler.SaveDesignAs
End Sub

Private Sub mnuViewOptions_Click()
    ShowOptions
End Sub

Private Sub txtWidth_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    If Len(txtWidth.Text) <> 0 And txtWidth <> "." Then
        If txtWidth.Text > 0 And txtWidth.Text <= (6 * 1440) - piRulerOffsetX Then
            clsHandler.CurrCtl.Width = txtWidth.Text * 1440
            clsHandler.ShowHandles False, False
            clsHandler.DesignChanged = True
        End If
    End If
End Sub

Public Sub txtXLeft_Change()

        
End Sub

Private Sub txtXLeft_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    bKeyMove = True
    If Not (clsHandler.CurrCtl Is Nothing) Then
        If Len(txtXLeft.Text) <> 0 And txtXLeft <> "." Then
            If txtXLeft.Text > 0 And txtXLeft.Text <= (6 * 1440) Then
                clsHandler.CurrCtl.Left = (txtXLeft.Text * 1440) + piRulerOffsetX
                clsHandler.ShowHandles False, False
                clsHandler.DesignChanged = True
            End If
        End If
    End If
    bKeyMove = False
End Sub

Private Sub txtYTop_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    bKeyMove = True
    If Not (clsHandler.CurrCtl Is Nothing) Then
        If Len(txtYTop.Text) <> 0 And txtYTop <> "." Then
            If txtYTop.Text > 0 And txtYTop.Text <= (2.8 * 1440) Then
                clsHandler.CurrCtl.Top = (txtYTop.Text * 1440) + piRulerOffsetY
                clsHandler.ShowHandles False, False
                clsHandler.DesignChanged = True
            End If
        End If
    End If
    bKeyMove = False
End Sub


Private Sub VRuler_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'clsHandler.HandlePositionIndicator -1, 0
    clsHandler.HandlePositionIndicator (VRuler.Left + x), (VRuler.Top + y)
End Sub

Private Sub PicHandle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    clsHandler.PicHandle_MouseDown Index
End Sub







