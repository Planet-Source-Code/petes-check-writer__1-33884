VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWriteCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarForm As frmCheckWrite 'local copy
Private mvarMode As enmMode 'local copy
Private m_objAutoCompleteUser As clsTextAutoComplete
Private mvarOpenCheckFile As String 'local copy
Public Sub SetOpenFile(Optional psCheckPath As String)
    
Dim cINi As New clsIniFile

    If FileExists(psCheckPath) Then
        If Len(psCheckPath) <> 0 Then
            mvarOpenCheckFile = psCheckPath
        End If
    
        cINi.FullPath = mvarOpenCheckFile
    
        mvarForm.Caption = "Write Check: " & cINi.ExtractName(mvarOpenCheckFile, exnFullName)
    
        Set cINi = Nothing
    End If
    
End Sub
Public Property Let OpenCheckFile(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OpenCheckFile = 5
    mvarOpenCheckFile = vData
End Property


Public Property Get OpenCheckFile() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OpenCheckFile
    OpenCheckFile = mvarOpenCheckFile
End Property



Public Sub LoadLogData()

Dim cINi As New clsIniFile
      
    If Not p_clsOptions.AutoComplete Then
        Exit Sub
    End If
    
    If Len(Trim$(mvarForm.txtPayTo.Text)) = 0 Then
        Exit Sub
    End If
    
    cINi.FullPath = App.path & "\checkwrite.log"
    cINi.SectionName = Trim$(mvarForm.txtPayTo.Text)
    mvarForm.txtAmount.Text = cINi.GetString("Amount", "")
    mvarForm.txtMemo.Text = cINi.GetString("Memo", "")
    mvarForm.txtAccountNo.Text = cINi.GetString("AccountNo", "")
    'cSave.PutString "Version", App.Major & "." & App.Minor & "." & App.Revision

    Set cINi = Nothing
        
    RefreshAutoComplete
    
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
        Set .CompleteTextbox = mvarForm.txtPayTo
        .Delimeter = vbNullChar
    End With
    
    Set cINi = Nothing
    
End Sub
Private Function Verified() As Boolean

Dim mfrmVerify As frmVerifyPrint
Dim itmX As ListItem

    Set mfrmVerify = New frmVerifyPrint
    Load mfrmVerify
    With mvarForm
         Set itmX = mfrmVerify.lvwVerify.ListItems.Add(1)
            itmX.Text = "Date:"
            itmX.SubItems(1) = .DTPicker.Value
            
        Set itmX = mfrmVerify.lvwVerify.ListItems.Add(2)
            itmX.Text = "Pay to the order of:"
            itmX.SubItems(1) = .txtPayTo.Text
        
        Set itmX = mfrmVerify.lvwVerify.ListItems.Add(3)
            itmX.Text = "Amount $"
            itmX.SubItems(1) = .txtAmount.Text
        
        Set itmX = mfrmVerify.lvwVerify.ListItems.Add(4)
            itmX.Text = ""
            itmX.SubItems(1) = .lblNumText.Caption
                  
        Set itmX = mfrmVerify.lvwVerify.ListItems.Add(5)
            itmX.Text = "Memo:"
            itmX.SubItems(1) = .txtMemo.Text
            
        Set itmX = mfrmVerify.lvwVerify.ListItems.Add(6)
            itmX.Text = "Account No:"
            itmX.SubItems(1) = .txtAccountNo.Text
            
        Set itmX = mfrmVerify.lvwVerify.ListItems.Add(7)
            itmX.Text = "Name:"
            itmX.SubItems(1) = .txtName.Text
    End With
    
    Set itmX = Nothing
    
    mfrmVerify.Show vbModal
    
    Verified = mfrmVerify.Verified
    
    Set mfrmVerify = Nothing
    
End Function
Private Sub LockFields(pbLock As Boolean)

    With mvarForm
        .txtAccountNo.Locked = pbLock
        .txtAmount.Locked = pbLock
        .txtMemo.Locked = pbLock
        .txtName.Locked = pbLock
        .txtPayTo.Locked = pbLock
        .DTPicker.Enabled = Not pbLock
    End With
    
    If pbLock Then
        With mvarForm
            .txtAccountNo.BackColor = &HC0C0C0
            .txtAmount.BackColor = &HC0C0C0
            .txtMemo.BackColor = &HC0C0C0
            .txtName.BackColor = &HC0C0C0
            .txtPayTo.BackColor = &HC0C0C0
        End With
    Else
        With mvarForm
            .txtAccountNo.BackColor = vbWhite
            .txtAmount.BackColor = vbWhite
            .txtMemo.BackColor = vbWhite
            .txtName.BackColor = vbWhite
            .txtPayTo.BackColor = vbWhite
        End With
    
    End If
    
    ClearFields
    
End Sub

Private Sub ClearFields()
    With mvarForm
        .txtAccountNo.Text = ""
        .txtAmount.Text = ""
        .txtMemo.Text = ""
        .txtName.Text = ""
        .txtPayTo.Text = ""
        .DTPicker.Value = Format(Now, "mm/dd/yyyy")
    End With
End Sub
Public Sub NewCheck()
    Me.Mode = md_New
    
    If Len(p_clsOptions.CheckWriterName) <> 0 Then
        mvarForm.txtName.Text = p_clsOptions.CheckWriterName
    End If
End Sub
Public Property Let Mode(ByVal vData As enmMode)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Mode = Form1
    mvarMode = vData
    
        Select Case mvarMode
            Case md_None '= 0
                With mvarForm
                    .mnuCheckNew.Enabled = True
                    .mnuCheckPrintCheck.Enabled = False
                End With
                
                LockFields True
                
            Case md_New '= 1
                With mvarForm
                    .mnuCheckNew.Enabled = True
                    .mnuCheckPrintCheck.Enabled = True
                End With
                
                LockFields False
            Case md_Edit '= 2

            Case md_View '= 3
                            
        End Select
        
        With mvarForm
            frmMain.tbWrite.Buttons("New").Enabled = .mnuCheckNew.Enabled
            frmMain.tbWrite.Buttons("Print").Enabled = .mnuCheckPrintCheck.Enabled
        End With
                
        SetMDIStatusBarMode mvarMode
End Property


Public Property Get Mode() As enmMode
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Mode
    Mode = mvarMode
End Property



Public Property Set Form(ByVal vData As frmCheckWrite)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Form = Form1
    Set mvarForm = vData
    
    mvarForm.Width = 6825
    mvarForm.Height = 3870
    
    RefreshAutoComplete
    
End Property


Public Property Get Form() As frmCheckWrite
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Form
    Set Form = mvarForm
    
    
End Property







Private Sub SaveToLog()

Dim cSave As New clsIniFile
 
    cSave.FullPath = App.path & "\checkwrite.log"

    cSave.SectionName = mvarForm.txtPayTo.Text
    cSave.PutString "Date", mvarForm.DTPicker.Value
    cSave.PutString "Amount", mvarForm.txtAmount.Text
    cSave.PutString "Memo", mvarForm.txtMemo.Text
    cSave.PutString "AccountNo", mvarForm.txtAccountNo.Text
    cSave.PutString "Version", App.Major & "." & App.Minor & "." & App.Revision
    Set cSave = Nothing
    
    RefreshAutoComplete
    
End Sub


Public Sub PrintData()

On Error GoTo PrintData_Error

    Dim cOpen As New clsIniFile
    Dim cOpenFile As New clsCmDlg
    Dim i As Integer
   
        
    If Len(mvarOpenCheckFile) = 0 Then
        cOpenFile.DefaultExt = "ckc"
        cOpenFile.DialogTitle = "Open check file"
        cOpenFile.FileName = ""
        cOpenFile.Filter = "Check Coordinates (*ckc)|*.ckc"
        cOpenFile.FilterIndex = 1
        cOpenFile.InitDir = App.path
        cOpenFile.ShowOpen
        Me.SetOpenFile cOpenFile.FileName
        
    End If
        
        If Len(mvarOpenCheckFile) <> 0 And Verified() Then
            ShowMessage anm_CopyToPrinter, "Sending check data to be printed", "Printing Check Data"
            
            
            If Not SelectPrinter() Then
                Unload frmShowAnim
                Exit Sub
            End If
            
            With cOpen
            
                If Right(mvarOpenCheckFile, 4) <> ".ckc" Then
                    mvarOpenCheckFile = mvarOpenCheckFile & ".ckc"
                End If
            
                .FullPath = mvarOpenCheckFile
            
                Printer.Font = "Courier New"
                Printer.FontSize = 10
                Printer.ScaleMode = vbTwips
                
                '[AccountInfo]
                .SectionName = "AccountInfo"
                Printer.CurrentX = .GetString("Left", 100) - piRulerOffsetX
                Printer.CurrentY = .GetString("Top", 100) - piRulerOffsetY
                Printer.Print mvarForm.txtAccountNo.Text
                
                '[Date]
                 .SectionName = "Date"
                Printer.CurrentX = .GetString("Left", 100) - piRulerOffsetX
                Printer.CurrentY = .GetString("Top", 100) - piRulerOffsetY
                Printer.Print Format(mvarForm.DTPicker.Value, "mm/dd/yyyy")
                
                '[PayTo]
                .SectionName = "PayTo"
                Printer.CurrentX = .GetString("Left", 100) - piRulerOffsetX
                Printer.CurrentY = .GetString("Top", 100) - piRulerOffsetY
                Printer.Print mvarForm.txtPayTo.Text
                
                '[AmountValue]
                Printer.FontSize = 12
                Printer.FontBold = True
                .SectionName = "AmountValue"
                Printer.CurrentX = .GetString("Left", 100) - piRulerOffsetX
                Printer.CurrentY = .GetString("Top", 100) - piRulerOffsetY
                Printer.Print mvarForm.txtAmount.Text & String(10 - Len(mvarForm.txtAmount.Text), "-")
                Printer.FontSize = 10
                Printer.FontBold = False
                
                '[AmountText]
                .SectionName = "AmountText"
                Printer.CurrentX = .GetString("Left", 100) - piRulerOffsetX
                Printer.CurrentY = .GetString("Top", 100) - piRulerOffsetY
                Printer.Print Mid(mvarForm.lblNumText.Caption, 1, 51) & String(51 - Len(Trim$(Mid(mvarForm.lblNumText.Caption, 1, 51))), "-") & IIf(Len(mvarForm.lblNumText.Caption) > 50, "...", "")
                
                '[Memo]
                .SectionName = "Memo"
                Printer.CurrentX = .GetString("Left", 100) - piRulerOffsetX
                Printer.CurrentY = .GetString("Top", 100) - piRulerOffsetY
                Printer.Print mvarForm.txtMemo.Text
                
                '[Name]
                .SectionName = "Name"
                Printer.CurrentX = .GetString("Left", 100) - piRulerOffsetX
                Printer.CurrentY = .GetString("Top", 100) - piRulerOffsetY
                Printer.Print mvarForm.txtName.Text
                
                'Me.OpenFile = .FullPath
                Printer.EndDoc
                SaveToLog
                Unload frmShowAnim
            End With
            
            Me.Mode = md_None
        End If
        
       
        Set cOpenFile = Nothing
        Set cOpen = Nothing
        
        
        
Exit Sub

PrintData_Error:
    
    If Not (cOpenFile Is Nothing) Then
        Set cOpenFile = Nothing
    End If
    
    If Not (cOpen Is Nothing) Then
        Set cOpen = Nothing
    End If
    
    Unload frmShowAnim
    
    MsgBox "Error #: " & str(Err) & vbCrLf & "Error Desc: " & Error & vbCrLf & "Sending check data to printer.", vbExclamation





End Sub

Private Sub Class_Terminate()
    If Not (mvarForm Is Nothing) Then
        Set mvarForm = Nothing
    End If
End Sub
