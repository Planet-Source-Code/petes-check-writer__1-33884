VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "SSA3D30.OCX"
Begin VB.Form frmNewBill 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Bill"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   ClipControls    =   0   'False
   Icon            =   "frmNewBill.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSCommand cmdSave 
      Height          =   660
      Left            =   2092
      TabIndex        =   9
      Top             =   1815
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   1164
      _Version        =   196609
      PictureFrames   =   1
      BackStyle       =   1
      Picture         =   "frmNewBill.frx":014A
      Caption         =   "Save "
      Alignment       =   4
      PictureAlignment=   1
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   90
      Picture         =   "frmNewBill.frx":0A24
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   8
      Top             =   825
      Width           =   960
   End
   Begin VB.TextBox txtPayTo 
      Height          =   300
      Left            =   2055
      TabIndex        =   0
      Top             =   180
      Width           =   4275
   End
   Begin VB.TextBox txtAmount 
      Height          =   300
      Left            =   2055
      TabIndex        =   2
      Top             =   1005
      Width           =   1530
   End
   Begin VB.TextBox txtLocation 
      Height          =   300
      Left            =   2055
      TabIndex        =   3
      Top             =   1410
      Width           =   3990
   End
   Begin MSComCtl2.DTPicker DTPicker 
      Height          =   315
      Left            =   2055
      TabIndex        =   1
      Top             =   585
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      _Version        =   393216
      Format          =   24510465
      CurrentDate     =   36827
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   660
      Left            =   3232
      TabIndex        =   10
      Top             =   1815
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   1164
      _Version        =   196609
      PictureFrames   =   1
      BackStyle       =   1
      Picture         =   "frmNewBill.frx":1599
      Caption         =   "Close "
      Alignment       =   4
      PictureAlignment=   1
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date:"
      Height          =   240
      Left            =   1140
      TabIndex        =   7
      Top             =   615
      Width           =   810
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Pay To The Order Of:"
      Height          =   480
      Left            =   1140
      TabIndex        =   6
      Top             =   135
      Width           =   900
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount: $"
      Height          =   240
      Left            =   1140
      TabIndex        =   5
      Top             =   1035
      Width           =   870
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Location:"
      Height          =   240
      Left            =   1140
      TabIndex        =   4
      Top             =   1425
      Width           =   1050
   End
End
Attribute VB_Name = "frmNewBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum enmBillTrackerMode
    btm_New = 2
    btm_Edit = 3
    btm_None = 0
    btm_Save = 1
End Enum

Private m_objAutoCompleteUser As clsTextAutoComplete
Public Event SaveData(psPayTo As String, psDueDate As String, psAmount As String, psLocation As String, pMode As enmBillTrackerMode)
Public FormMode As enmBillTrackerMode




Public Sub InitForm(pMode As enmBillTrackerMode)
    
    FormMode = pMode
    
    If FormMode = btm_New Then
        Me.Icon = LoadResPicture("NEW", vbResIcon)
        Me.Caption = "New Bill"
    Else
        Me.Icon = LoadResPicture("EDIT", vbResIcon)
        Me.Caption = "Edit Bill"
    End If
    
    
    
    DTPicker.Value = Format(Now, "mm/dd/yyyy")
    RefreshAutoComplete
    
End Sub

Private Sub cmdExit_Click()
    FormMode = btm_None
    RaiseEvent SaveData("", "", "", "", FormMode)
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
    FormMode = btm_Save
    RaiseEvent SaveData(txtPayTo.Text, Format(DTPicker.Value, "mm/dd/yyyy"), txtAmount.Text, txtLocation.Text, FormMode)
    Unload Me
End Sub

Private Sub DTPicker_KeyPress(KeyAscii As Integer)
    If p_clsOptions.EnterTab And KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If p_clsOptions.EnterTab And KeyCode = 13 Then
        SendKeys "{TAB}"
        KeyCode = 0
    End If
End Sub
Private Sub RefreshAutoComplete()
    Dim cINi As New clsIniFile
    
    If Not p_clsOptions.AutoComplete Then
        Exit Sub
    End If
    
    If Not (m_objAutoCompleteUser Is Nothing) Then
        Set m_objAutoCompleteUser = Nothing
    End If
    
    Set m_objAutoCompleteUser = New clsTextAutoComplete
    
    cINi.FullPath = App.path & "\checkwrite.log"
    
    
    With m_objAutoCompleteUser
        .SearchList = cINi.GetSections
        Set .CompleteTextbox = txtPayTo
        .Delimeter = vbNullChar
    End With
    
    Set cINi = Nothing
    
End Sub
Private Sub LoadLogData()
Dim cINi As New clsIniFile
      
    If Not p_clsOptions.AutoComplete Then
        Exit Sub
    End If
    
    If Len(Trim$(mvarForm.txtPayTo.Text)) = 0 Then
        Exit Sub
    End If
    
    cINi.FullPath = App.path & "\checkwrite.log"
    cINi.SectionName = Trim$(mvarForm.txtPayTo.Text)
    txtAmount.Text = cINi.GetString("Amount", "")
    txtMemo.Text = cINi.GetString("Memo", "")
    txtAccountNo.Text = cINi.GetString("AccountNo", "")
    'cSave.PutString "Version", App.Major & "." & App.Minor & "." & App.Revision

    Set cINi = Nothing
        
    RefreshAutoComplete
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not (m_objAutoCompleteUser Is Nothing) Then
        Set m_objAutoCompleteUser = Nothing
    End If
    
    
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    If p_clsOptions.EnterTab And KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub



Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    If p_clsOptions.EnterTab And KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPayTo_KeyPress(KeyAscii As Integer)
    If p_clsOptions.EnterTab And KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub




Public Property Let DueDate(ByVal psDueDate As String)
    On Error Resume Next
    DTPicker.Value = psDueDate
End Property

Public Property Let Amount(ByVal psAmount As String)
    txtAmount.Text = psAmount
End Property

Public Property Let PayTo(ByVal psPayTo As String)
    txtPayTo.Text = psPayTo
End Property
Public Property Let Location(ByVal psLocation As String)
    txtLocation.Text = psLocation
End Property
