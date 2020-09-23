VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Check Writer"
   ClientHeight    =   6420
   ClientLeft      =   435
   ClientTop       =   1005
   ClientWidth     =   9540
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   660
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   582
      ButtonWidth     =   1561
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Write "
            Key             =   "Write"
            Object.ToolTipText     =   "Write a check"
            ImageKey        =   "WriteCheck"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Design"
            Key             =   "Design"
            Object.ToolTipText     =   "Design a check"
            ImageKey        =   "Design"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print   "
            Key             =   "Print"
            Object.ToolTipText     =   "Print Selection"
            ImageKey        =   "Printer"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Holder"
                  Text            =   "Holder"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Sample"
                  Text            =   "Sample"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit    "
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit/End Program"
            ImageKey        =   "Exit"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   8340
      Top             =   1155
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A26
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B82
            Key             =   "View"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CDE
            Key             =   "None"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E3A
            Key             =   "Design"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F9A
            Key             =   "Print Holder"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1536
            Key             =   "PrintSample"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AD2
            Key             =   "WriteCheck"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":206E
            Key             =   "Printer"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":260A
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2926
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A82
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2BDE
            Key             =   "Close"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   6120
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   6359
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2725
            Picture         =   "frmMain.frx":2EFA
            Text            =   " None"
            TextSave        =   " None"
            Key             =   "sbMode"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "03/10/2002"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "1:10 PM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   706
            MinWidth        =   706
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbDesign 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   330
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   582
      ButtonWidth     =   1402
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "New"
            Object.ToolTipText     =   "New Design"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "Open"
            Object.ToolTipText     =   "Open check design"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
            Object.ToolTipText     =   "Save Design"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit  "
            Key             =   "Edit"
            Object.ToolTipText     =   "Edit/Design"
            ImageKey        =   "Edit"
            Style           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Printer"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Design"
                  Text            =   "Design"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Holder"
                  Text            =   "Holder"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Sample"
                  Text            =   "Sample"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Key             =   "Close"
            Object.ToolTipText     =   "Close Design Screen"
            ImageKey        =   "Close"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbWrite 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   582
      ButtonWidth     =   1561
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New  "
            Key             =   "New"
            Object.ToolTipText     =   "New Design"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Design"
            Key             =   "Design"
            Object.ToolTipText     =   "Design check"
            ImageKey        =   "Design"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print  "
            Key             =   "Print"
            Object.ToolTipText     =   "Print Selection"
            ImageKey        =   "Printer"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Data"
                  Text            =   "Data"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Holder"
                  Text            =   "Holder"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Sample"
                  Text            =   "Sample"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close "
            Key             =   "Close"
            Object.ToolTipText     =   "Close"
            ImageKey        =   "Close"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuCheck 
      Caption         =   "&Check"
      Begin VB.Menu mnuCheckDesign 
         Caption         =   "&Design"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuCheckPrint 
         Caption         =   "&Print"
         Begin VB.Menu mnuCheckPrintHolder 
            Caption         =   "&Check Holder"
         End
         Begin VB.Menu mnuCheckPrintSample 
            Caption         =   "&Sample Check"
         End
      End
      Begin VB.Menu mnuCheckWrite 
         Caption         =   "&Write"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuCheckSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpContents_Click()
    Call HtmlHelp(0, App.HelpFile, HH_DISPLAY_TOPIC, ByVal 0&)
End Sub

Private Sub mnuViewWizards_Click()
Load frmNavigator
    frmNavigator.InitForm wiz_FirstTime
    frmNavigator.Show , Me
End Sub

Private Sub mnuWindowArrangeIcons_Click()
  Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowCascade_Click()
  Me.Arrange vbCascade
End Sub



Private Sub mnuWindowTileHorizontal_Click()
  Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowTileVertical_Click()
  Me.Arrange vbTileVertical
End Sub


Private Sub MDIForm_Activate()
    'ShowToolbar tb_Main
    
End Sub

Private Sub MDIForm_Load()
    'frmShareware.Show vbModal
    
    Me.Caption = "Check Writer - Version " & App.Major & "." & App.Minor & "." & App.Revision
    ShowToolbar tb_Main
    Set p_clsOptions = New clsOptions
    Set p_clsOptions.Form = frmOptions
    p_clsOptions.LoadOptions
    SetFormPosition Me, False
    
    If p_clsOptions.ShowWizards Then
        Load frmNavigator
        frmNavigator.InitForm wiz_FirstTime
        frmNavigator.Show , Me
    End If
    
   
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    
    
    If Not (p_clsOptions Is Nothing) Then
        Set p_clsOptions = Nothing
    End If
    
    If Me.WindowState <> vbMaximized Then
        SetFormPosition Me, True
    End If
    
End Sub

Private Sub mnuCheckDesign_Click()
    Load frmCheckRuler
  frmCheckRuler.InitForm
  frmCheckRuler.Show
End Sub

Private Sub mnuCheckExit_Click()
    Unload Me
End Sub

Private Sub mnuCheckPrintHolder_Click()
   
   PrintHolder
   
    
End Sub

Private Sub mnuCheckPrintSample_Click()
     PrintHolder True
End Sub

Private Sub mnuCheckWrite_Click()
    Load frmCheckWrite
    frmCheckWrite.InitForm
    frmCheckWrite.Show
End Sub

Private Sub mnuViewOptions_Click()
    ShowOptions
End Sub

Private Sub tbDesign_ButtonClick(ByVal Button As MSComctlLib.Button)
    With frmCheckRuler
        Select Case Button.Key
        
            Case "New"
                .mnuDesignNew_Click
            Case "Open"
                .mnuDesignOpen_Click
            Case "Print"
               ShowPrintCheck False, True
            Case "Edit"
                .mnuDesignDesign_Click
            Case "Save"
                .mnuDesignSave_Click
            Case "Close"
                Unload frmCheckRuler
        End Select
    End With
End Sub

Private Sub tbDesign_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Holder"
            mnuCheckPrintHolder_Click
        Case "Sample"
            mnuCheckPrintSample_Click
        Case "Design"
            frmCheckRuler.mnuDesignPrintDesignSample_Click
    End Select
End Sub

Private Sub tbWrite_ButtonClick(ByVal Button As MSComctlLib.Button)
    With frmCheckWrite
        Select Case Button.Key
        
            Case "New"
                .mnuCheckNew_Click
            Case "Design"
                mnuCheckDesign_Click
            Case "Print"
               ShowPrintCheck True, False
            Case "Close"
                Unload frmCheckWrite
        End Select
    End With
End Sub

Private Sub tbWrite_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Holder"
            mnuCheckPrintHolder_Click
        Case "Sample"
            mnuCheckPrintSample_Click
        Case "Data"
            frmCheckWrite.mnuCheckPrintCheck_Click
    End Select
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    
        Case "Write"
            mnuCheckWrite_Click
        Case "Design"
            mnuCheckDesign_Click
        Case "Print"
            ShowPrintCheck
        Case "Exit"
            mnuCheckExit_Click
    End Select
End Sub

Private Sub tlbMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Holder"
            mnuCheckPrintHolder_Click
        Case "Sample"
            mnuCheckPrintSample_Click
    End Select
    
End Sub
