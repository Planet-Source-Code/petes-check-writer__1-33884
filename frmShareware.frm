VERSION 5.00
Begin VB.Form frmShareware 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Check Writer Shareware Version"
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Height          =   330
      Left            =   5295
      TabIndex        =   3
      Top             =   1170
      Width           =   1290
   End
   Begin VB.Timer tmrColor 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3705
      Top             =   540
   End
   Begin VB.Timer tmrUnload 
      Interval        =   1000
      Left            =   2775
      Top             =   510
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H8000000D&
      Caption         =   " Check Writer Shareware Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6720
   End
   Begin VB.Label lblUses 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   30
      TabIndex        =   1
      Top             =   1635
      Width           =   6585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmShareware.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   6585
   End
End
Attribute VB_Name = "frmShareware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lUnloadTime As Long


Private Sub cmdRegister_Click()
    LoadWebPage "http://pjs-inc.com/CheckWriterSoftware/Register.htm", Me
End Sub

Private Sub Form_Load()

Dim lUses As Long

    lUses = Val(GetSetting(App.ProductName, "Uses", "Count"))
    SaveSetting App.ProductName, "Uses", "Count", Trim$(str$(lUses + 1))
    
    If lUses <= 1 Then
        m_lUnloadTime = 30000
    Else
        m_lUnloadTime = 10000
    End If
    
    If lUses < 0 Then
        lUses = 30
    End If
    
    If lUses + 1 >= 30 Then
        tmrColor.Enabled = True
    End If
    
    lblUses.Caption = "You have used Check Writer: " & Trim$(str(lUses + 1)) & " times!"
    
    
End Sub

Private Sub tmrColor_Timer()
Static lColor As Long

    If lColor = 0 Then
        lColor = vbRed
    End If
    
    If lColor = vbRed Then
        lColor = vbYellow
    Else
        lColor = vbRed
    End If
    
    frmShareware.BackColor = lColor
    
    
End Sub

Private Sub tmrUnload_Timer()
Static Count As Integer

    Count = Count + 1000
    
    lblTitle.Caption = " Check Writer Shareware >> Loading in: " & Trim$(str$((m_lUnloadTime - Count) / 1000)) & " seconds"
    If Count = m_lUnloadTime Then
        MsgBox "Registering will remove this nagging screen as well as this message box!", vbExclamation, "Please Register!"
        Unload Me
    End If
End Sub
