VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmShowAnim 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please Wait..."
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   195
      Left            =   30
      TabIndex        =   3
      Top             =   1785
      Visible         =   0   'False
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   1725
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.Animation animAVI 
      Height          =   975
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   1720
      _Version        =   393216
      AutoPlay        =   -1  'True
      Center          =   -1  'True
      FullWidth       =   347
      FullHeight      =   65
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   45
      TabIndex        =   1
      Top             =   1020
      Width           =   5220
   End
End
Attribute VB_Name = "frmShowAnim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Public Sub Parameters(piAVI_Type As enmAnimations, Optional psMsg As String, Optional psCaption As String)

On Error Resume Next



    
    
    If Len(psCaption) <> 0 Then
        Me.Caption = psCaption
    Else
        Me.Caption = "Processing..."
    End If
    
    If Len(psMsg) <> 0 Then
        lblMessage.Caption = psMsg
    Else
        lblMessage.Caption = "Please Wait..."
    End If
    
    ClearAnim animAVI
    LoadResAVI animAVI, piAVI_Type
    Me.Refresh
    
End Sub

Public Sub Progress(piValue As Variant, piTotal As Variant)
On Error Resume Next
    If piValue <> 0 Then
        ProgressBar1.Visible = True
        StatusBar1.Visible = True
    Else
        ProgressBar1.Visible = False
        StatusBar1.Visible = False
        Exit Sub
    End If
    
    ProgressBar1.Max = piTotal
    ProgressBar1.Value = piValue
    StatusBar1.Panels(2).Text = Trim$(str(piValue)) & " of " & Trim$(str(piTotal)) & " - " & Trim$(str$(CInt((piValue / piTotal) * 100))) & "%"
    
    If ProgressBar1.Max = ProgressBar1.Value Then
        ProgressBar1.Value = 0
        ProgressBar1.Visible = False
        StatusBar1.Visible = False
    End If
    'Me.Refresh
    
    
    
End Sub


