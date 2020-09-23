VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarAutoComplete As Boolean 'local copy
Private mvarMoveFieldSize As String 'local copy
Private mvarDefaultCheckPath As String 'local copy
Private mvarDefaultPrinter As String 'local copy
Private mvarEditOnOpen As Boolean 'local copy
Private mvarForm As frmOptions 'local copy
'local variable(s) to hold property value(s)
Private mvarCheckWriterName As String 'local copy
'local variable(s) to hold property value(s)
Private mvarShowWizards As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarEnterTab As Boolean 'local copy
Public Property Let EnterTab(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EnterTab = 5
    mvarEnterTab = vData
End Property


Public Property Get EnterTab() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EnterTab
    EnterTab = mvarEnterTab
End Property



Public Property Let ShowWizards(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ShowWizards = 5
    mvarShowWizards = vData
End Property


Public Property Get ShowWizards() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ShowWizards
    ShowWizards = mvarShowWizards
End Property



Public Property Let CheckWriterName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CheckWriterName = 5
    mvarCheckWriterName = vData
End Property


Public Property Get CheckWriterName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CheckWriterName
    CheckWriterName = mvarCheckWriterName
End Property



Public Function SelectDesign() As String
   
 
    Dim cOpenFile As New clsCmDlg
    Dim i As Integer
    
      
        
        cOpenFile.DefaultExt = "ckc"
        cOpenFile.DialogTitle = "Open check file"
        cOpenFile.FileName = ""
        cOpenFile.Filter = "Check Coordinates (*ckc)|*.ckc"
        cOpenFile.FilterIndex = 1
        cOpenFile.InitDir = App.path
        cOpenFile.ShowOpen
        
        If Len(cOpenFile.FileName) <> 0 Then
           If Right(cOpenFile.FileName, 4) <> ".ckc" Then
                cOpenFile.FileName = cOpenFile.FileName & ".ckc"
           End If
        End If
        
        SelectDesign = cOpenFile.FileName
        
        Set cOpenFile = Nothing
        

End Function

Public Function SelectPrinter() As String

On Error GoTo SelectPrinter_Error

Dim cPrinter As New clsCmDlg
    
    cPrinter.flags = cdlPDDisablePrintToFile + cdlPDNoPageNums + cdlPDNoSelection
    cPrinter.CancelError = True
    cPrinter.ShowPrinter
    
    
    mvarDefaultPrinter = Printer.DeviceName
    SelectPrinter = mvarDefaultPrinter
    
    Me.SetDefaultPrinter
    
 Exit Function
SelectPrinter_Error:
    If Err = 5 Then
        SelectPrinter = ""
    End If
    
    Set cPrinter = Nothing
End Function
Public Sub SetDefaultPrinter(Optional psDefaultPrinter As String)

 Dim objPrinter As Printer
 Dim sDefaultPrinter As String
 
    If Len(psDefaultPrinter) <> 0 Then
        sDefaultPrinter = psDefaultPrinter
    Else
        sDefaultPrinter = mvarDefaultPrinter
    End If
    
    If Len(sDefaultPrinter) <> 0 Then
        If Printer.DeviceName <> sDefaultPrinter Then
            For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = UCase(sDefaultPrinter) Then
                    Set Printer = objPrinter
                    mvarDefaultPrinter = sDefaultPrinter
                End If
            Next
        End If
    End If
    
    
End Sub
Public Property Set Form(ByVal vData As frmOptions)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Form = Form1
    Set mvarForm = vData
End Property


Public Property Get Form() As frmOptions
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Form
    Set Form = mvarForm
End Property



Public Property Let EditOnOpen(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EditOnOpen = 5
    mvarEditOnOpen = vData
End Property


Public Property Get EditOnOpen() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EditOnOpen
    EditOnOpen = mvarEditOnOpen
End Property
Public Sub Apply()
    
    
    With mvarForm
        mvarAutoComplete = IIf(.chkAutoComplete.Value = vbChecked, True, False)
        mvarMoveFieldSize = IIf(Len(.txtMoveIncrement.Text) <> 0, .txtMoveIncrement.Text, "0.10")
        mvarDefaultCheckPath = Trim$(.txtDefaultDesign.Text)
        mvarDefaultPrinter = Trim$(.txtDefaultPrinter.Text)
        mvarEditOnOpen = IIf(.chkEditOnOpen.Value = vbChecked, True, False)
        mvarCheckWriterName = Trim$(.txtFullName.Text)
        mvarEnterTab = IIf(.chkEnterTab.Value = vbChecked, True, False)
    End With
    
    
    Me.SaveOptions
    
End Sub


Public Property Let DefaultPrinter(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DefaultPrinter = 5
    mvarDefaultPrinter = UCase(vData)
End Property


Public Property Get DefaultPrinter() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DefaultPrinter
    DefaultPrinter = mvarDefaultPrinter
End Property



Public Property Let DefaultCheckPath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DefaultCheckPath = 5
    mvarDefaultCheckPath = vData
End Property


Public Property Get DefaultCheckPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DefaultCheckPath

    If InStr(1, mvarDefaultCheckPath, "\") = 0 Then
        mvarDefaultCheckPath = App.path & "\" & mvarDefaultCheckPath
        mvarDefaultCheckPath = Replace(mvarDefaultCheckPath, "\\", "\")
    End If
    
    DefaultCheckPath = mvarDefaultCheckPath
End Property



Public Property Let MoveFieldSize(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MoveFieldSize = 5
    mvarMoveFieldSize = vData
End Property


Public Property Get MoveFieldSize() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MoveFieldSize
    MoveFieldSize = mvarMoveFieldSize
End Property



Public Sub SaveOptions()
    SaveSetting App.ProductName, "Options", "AutoComplete", mvarAutoComplete
    SaveSetting App.ProductName, "Options", "MoveFieldSize", mvarMoveFieldSize
    SaveSetting App.ProductName, "Options", "DefaultDesign", mvarDefaultCheckPath
    SaveSetting App.ProductName, "Options", "DefaultPrinter", mvarDefaultPrinter
    SaveSetting App.ProductName, "Options", "EditOnOpen", mvarEditOnOpen
    SaveSetting App.ProductName, "Options", "CheckWriterName", mvarCheckWriterName
    SaveSetting App.ProductName, "Options", "ShowWizards", mvarShowWizards
    SaveSetting App.ProductName, "Options", "EnterTab", mvarEnterTab
End Sub

Public Sub LoadOptions()
   mvarAutoComplete = CBool(GetSetting(App.ProductName, "Options", "AutoComplete", "False"))
   mvarMoveFieldSize = GetSetting(App.ProductName, "Options", "MoveFieldSize", "0.10")
   mvarDefaultCheckPath = GetSetting(App.ProductName, "Options", "DefaultDesign", "")
   mvarDefaultPrinter = GetSetting(App.ProductName, "Options", "DefaultPrinter", "")
   mvarEditOnOpen = CBool(GetSetting(App.ProductName, "Options", "EditOnOpen", "False"))
   mvarCheckWriterName = GetSetting(App.ProductName, "Options", "CheckWriterName", "")
   mvarShowWizards = CBool(GetSetting(App.ProductName, "Options", "ShowWizards", "True"))
    mvarEnterTab = CBool(GetSetting(App.ProductName, "Options", "EnterTab", "False"))
    
   With mvarForm
        .chkAutoComplete.Value = IIf(mvarAutoComplete = True, vbChecked, vbUnchecked)
        .txtMoveIncrement.Text = mvarMoveFieldSize
        .txtDefaultDesign.Text = mvarDefaultCheckPath
        .txtDefaultPrinter.Text = mvarDefaultPrinter
        .chkEditOnOpen.Value = IIf(mvarEditOnOpen = True, vbChecked, vbUnchecked)
        .txtFullName.Text = mvarCheckWriterName
        .chkEnterTab = IIf(mvarEnterTab = True, vbChecked, vbUnchecked)
    End With
End Sub

Public Property Let AutoComplete(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AutoComplete = 5
    mvarAutoComplete = vData
End Property


Public Property Get AutoComplete() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AutoComplete
    AutoComplete = mvarAutoComplete
End Property




Private Sub Class_Terminate()
    If Not (mvarForm Is Nothing) Then
        Unload frmOptions
        Set mvarForm = Nothing
    End If
End Sub
