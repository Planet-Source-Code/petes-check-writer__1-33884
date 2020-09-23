VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNavigator 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome to Check Writer"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "frmNavigator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5430
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList imlIcons16 
      Left            =   4755
      Top             =   1695
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavigator.frx":058A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavigator.frx":06E6
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavigator.frx":0842
            Key             =   "View"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavigator.frx":099E
            Key             =   "None"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavigator.frx":0AFA
            Key             =   "Design"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavigator.frx":0C5A
            Key             =   "Print Holder"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavigator.frx":11F6
            Key             =   "PrintSample"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavigator.frx":1792
            Key             =   "WriteCheck"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavigator.frx":1D2E
            Key             =   "Printer"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavigator.frx":22CA
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavigator.frx":25E6
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavigator.frx":2742
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavigator.frx":289E
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavigator.frx":2BBA
            Key             =   "Wizard"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNavigator.frx":3154
            Key             =   "Options"
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkShow 
      BackColor       =   &H80000005&
      Caption         =   "Do not show on start up"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   7
      Top             =   3225
      Width           =   2145
   End
   Begin MSComctlLib.ListView lvwList 
      Height          =   2865
      Left            =   945
      TabIndex        =   6
      Top             =   270
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   5054
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      SmallIcons      =   "imlIcons16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
      Picture         =   "frmNavigator.frx":36EE
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   3495
      Left            =   0
      ScaleHeight     =   3435
      ScaleWidth      =   885
      TabIndex        =   8
      Top             =   0
      Width           =   945
      Begin VB.OptionButton optOptions 
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   30
         Picture         =   "frmNavigator.frx":5637
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Show option wizard"
         Top             =   2265
         Width           =   810
      End
      Begin VB.OptionButton optFirstTime 
         Caption         =   "First Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   30
         Picture         =   "frmNavigator.frx":5BC1
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Show list of things to do first"
         Top             =   45
         Value           =   -1  'True
         Width           =   810
      End
      Begin VB.OptionButton optWrite 
         Caption         =   "Write"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   30
         Picture         =   "frmNavigator.frx":5D0B
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Show check writing wizards"
         Top             =   1710
         Width           =   810
      End
      Begin VB.OptionButton optPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   30
         Picture         =   "frmNavigator.frx":6295
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Show print wizards"
         Top             =   1155
         Width           =   810
      End
      Begin VB.OptionButton optDesign 
         Caption         =   "Design"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   30
         Picture         =   "frmNavigator.frx":681F
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Show design wizards"
         Top             =   600
         Width           =   810
      End
      Begin VB.OptionButton optExit 
         Height          =   570
         Left            =   30
         Picture         =   "frmNavigator.frx":6BC1
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Close"
         Top             =   2820
         Width           =   810
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000011&
      Caption         =   " Check Writer Wizards...."
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
      Height          =   270
      Left            =   945
      TabIndex        =   9
      Top             =   0
      Width           =   4530
   End
End
Attribute VB_Name = "frmNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cWiz As clsWizard
Public Sub InitForm(pWizard As enmWizardTypes)

    Set cWiz = New clsWizard
    Set cWiz.NavForm = Me
    cWiz.WizardType = pWizard
    
    If p_clsOptions.ShowWizards = False Then
        chkShow.Value = vbChecked
    End If
    
    Select Case pWizard
        Case wiz_Design '= 1
            optDesign.Value = True
        Case wiz_Print '= 2
            optPrint.Value = True
        Case wiz_Write '= 3
            optWrite.Value = True
        Case wiz_FirstTime '= 0
            optFirstTime.Value = True
        Case wiz_Option '= 4
            optOptions.Value = True
    End Select
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cWiz = Nothing
    
    If chkShow.Value = vbChecked Then
        p_clsOptions.ShowWizards = False
        p_clsOptions.SaveOptions
    Else
        p_clsOptions.ShowWizards = True
        p_clsOptions.SaveOptions
    End If
    
End Sub

Private Sub lvwList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Select Case Item.Key
        Case "New1"
            frmDesignWizard.Show , Me
        Case "Write1"
            frmWriteWizard.Show , Me
        Case "Option1"
            frmOptions.Show , Me
        Case "Sample1"
            PrintHolder True
            MsgBox "After the check sample prints..." & vbCrLf & "Put the sample check back into the printer so it prints on the same side as the sample check." & vbCrLf & "Click OK when ready!", vbInformation, "Print Sample Data"
            PrintSampleData
        Case "Print1"
            Load frmCheckRuler
            frmCheckRuler.InitForm
            frmCheckRuler.clsHandler.OpenDesign
            frmCheckRuler.mnuDesignPrintDesignSample_Click
            Unload frmCheckRuler
        Case "Print2"
            PrintHolder
        Case "Print3"
            PrintHolder True
        Case "Print4"
             MsgBox "Put the sample check into the printer so it prints on the same side as the sample check." & vbCrLf & "Click OK when ready!", vbInformation, "Print Sample Data"
            PrintSampleData
    End Select
End Sub

Private Sub optDesign_Click()
    cWiz.WizardType = wiz_Design
End Sub

Private Sub optExit_Click()
    Unload Me
End Sub

Private Sub optFirstTime_Click()
    cWiz.WizardType = wiz_FirstTime
End Sub

Private Sub optOptions_Click()
    cWiz.WizardType = wiz_Option
End Sub

Private Sub optPrint_Click()
    cWiz.WizardType = wiz_Print
End Sub

Private Sub optWrite_Click()
    cWiz.WizardType = wiz_Write
End Sub
