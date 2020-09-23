VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBillTracker 
   Caption         =   "Bill Tracker"
   ClientHeight    =   4650
   ClientLeft      =   165
   ClientTop       =   630
   ClientWidth     =   9000
   Icon            =   "frmBillTracker.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4650
   ScaleWidth      =   9000
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   4080
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   1005
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2170
            MinWidth        =   882
            Picture         =   "frmBillTracker.frx":030A
            Text            =   " Total "
            TextSave        =   " Total "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   13176
            Text            =   "$0.00"
            TextSave        =   "$0.00"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lvwBills 
      Height          =   4110
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   7250
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Due Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Pay To"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Location"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuBill 
      Caption         =   "&Bill"
      Begin VB.Menu mnuBillNew 
         Caption         =   "&New"
         Begin VB.Menu mnuBillNewSingle 
            Caption         =   "&Single"
         End
         Begin VB.Menu mnuBillNewRecur 
            Caption         =   "&Recurring"
         End
      End
      Begin VB.Menu mnuBillEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuBillDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuBillPrint 
         Caption         =   "&Print"
         Begin VB.Menu mnuBillPrintCheck 
            Caption         =   "&Check"
         End
         Begin VB.Menu mnuBillPrintList 
            Caption         =   "&List"
         End
      End
      Begin VB.Menu mnuBillSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBillClose 
         Caption         =   "C&lose"
      End
   End
   Begin VB.Menu mnuMaint 
      Caption         =   "&Maintenance"
      Begin VB.Menu mnuMaintProcess 
         Caption         =   "&Process Paid"
      End
      Begin VB.Menu mnuMaintRecur 
         Caption         =   "&Recurring Bills"
      End
   End
End
Attribute VB_Name = "frmBillTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents clsBills As clsBillTracker
Attribute clsBills.VB_VarHelpID = -1

Public Sub InitForm()

    Set clsBills = New clsBillTracker
    
    clsBills.DataFile = App.path & "\Bills.xml"
    clsBills.LogFile = App.path & "\Paid.xml"
    
    Call clsBills_RefreshGrid("")
    
    Me.Width = 9120
    Me.Height = 5340
End Sub

Private Sub clsBills_RefreshGrid(psXMLFile As String)
Dim i As Long
Dim lTotal As Currency

    clsBills.FillListView lvwBills, clsBills.DataFile
    
    For i = 1 To lvwBills.ListItems.Count
        lTotal = lTotal + lvwBills.ListItems(i).SubItems(3)
    Next i
    
    sbStatusBar.Panels(2) = Format(lTotal, "$#,#.#0")
    
End Sub

Private Sub Form_Load()
    InitForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set clsBills = Nothing
End Sub

Private Sub lvwBills_DblClick()
    mnuBillEdit_Click
End Sub

Private Sub mnuBillClose_Click()
    Unload Me
End Sub

Private Sub mnuBillDelete_Click()
    If Not lvwBills.SelectedItem Is Nothing Then
        clsBills.DeleteBill lvwBills.SelectedItem.Index
    End If
    
End Sub

Private Sub mnuBillEdit_Click()
    If Not lvwBills.SelectedItem Is Nothing Then
        clsBills.EditBill lvwBills.SelectedItem.Index
    End If
End Sub

Private Sub mnuBillNewSingle_Click()
    clsBills.AddBill
End Sub
