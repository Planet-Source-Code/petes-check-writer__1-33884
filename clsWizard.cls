VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarNavForm As frmNavigator 'local copy

Public Enum enmWizardTypes
    wiz_Design = 1
    wiz_Print = 2
    wiz_Write = 3
    wiz_FirstTime = 0
    wiz_Option = 4
End Enum
'local variable(s) to hold property value(s)
Private mvarWizardType As enmWizardTypes 'local copy
Private Sub FillNavigator()


Dim itmX As ListItem

    If mvarNavForm Is Nothing Then
        Exit Sub
    End If
    
    With mvarNavForm.lvwList
        .ListItems.Clear
        
    Select Case mvarWizardType
        Case wiz_Design '= 1
            'colWizards.Add "New check design"
           Set itmX = .ListItems.Add("1", "New1", "Design your check", , "Design")
        Case wiz_Print '= 2
            'colWizards.Add "Check design sample"
            Set itmX = .ListItems.Add("1", "Print1", "Print your design", , "Printer")
            'colWizards.Add "Check holder"
            Set itmX = .ListItems.Add("2", "Print2", "Print a check holder", , "Print Holder")
            'colWizards.Add "Check sample"
            Set itmX = .ListItems.Add("3", "Print3", "Print a check sample", , "PrintSample")
            Set itmX = .ListItems.Add("4", "Print4", "Print sample data on a sample check", , "PrintSample")
        Case wiz_Write '= 3
            'colWizards.Add "Write a check"
            Set itmX = .ListItems.Add("1", "Write1", "Write a check", , "WriteCheck")
        Case wiz_Option
            Set itmX = .ListItems.Add("1", "Option1", "Set your preferences & options", , "Options")
        Case wiz_FirstTime
            Set itmX = .ListItems.Add("1", "Sample1", "1. Print a sample check", , "PrintSample")
            Set itmX = .ListItems.Add("2", "New1", "2. Design your check", , "Design")
            Set itmX = .ListItems.Add("3", "Print2", "3. Print a check holder", , "Print Holder")
            Set itmX = .ListItems.Add("4", "Write1", "4. Write and print your check data", , "WriteCheck")
            Set itmX = .ListItems.Add("5", "Option1", "5. Set your preferences & options", , "Options")
    End Select
    
    
    If Not (itmX Is Nothing) Then
        Set itmX = Nothing
    End If
        
    End With
End Sub
Public Property Let WizardType(ByVal vData As enmWizardTypes)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WizardType = 5
    mvarWizardType = vData
    
    FillNavigator
End Property


Public Property Get WizardType() As enmWizardTypes
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WizardType
    WizardType = mvarWizardType
End Property




Public Property Set NavForm(ByVal vData As frmNavigator)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.NavForm = Form1
    Set mvarNavForm = vData
    Me.WizardType = wiz_FirstTime
End Property


Public Property Get NavForm() As frmNavigator
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NavForm
    Set NavForm = mvarNavForm
    
End Property



Private Sub Class_Terminate()
    If Not (mvarNavForm Is Nothing) Then
        Set mvarNavForm = Nothing
    End If
End Sub
