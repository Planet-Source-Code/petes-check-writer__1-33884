VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBillTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarDataFile As String 'local copy
Private mvarLogFile As String 'local copy
Private WithEvents mvarNewBill As frmNewBill
Attribute mvarNewBill.VB_VarHelpID = -1
Private m_bNewDataSave As Boolean

Private Type BillData
    DueDate As String
    Amount As String
    Location As String
    PayTo As String
    Mode As enmBillTrackerMode
End Type

Private m_NewBillData As BillData

Public Event RefreshGrid(psXMLFile As String)
Public Sub EditBill(plID As Long)
    Dim objDom As DOMDocument
    Dim objRootElem As IXMLDOMElement
    
    Dim objCurNode As IXMLDOMNode
    Dim objDelNode As IXMLDOMNode
    
   
    Dim j As Long
    
    Set objDom = New DOMDocument
    
    Set mvarNewBill = New frmNewBill
    Load mvarNewBill
    
    If objDom.Load(Me.DataFile) Then
        Set objRootElem = objDom.documentElement
        For j = 0 To objRootElem.childNodes.Item(plID - 1).childNodes.length - 1
            Select Case UCase(objRootElem.childNodes.Item(plID - 1).childNodes(j).nodeName)
                Case "DUEDATE"
                    mvarNewBill.DueDate = objRootElem.childNodes.Item(plID - 1).childNodes(j).Text
                Case "AMOUNT"
                    mvarNewBill.Amount = objRootElem.childNodes.Item(plID - 1).childNodes(j).Text
                Case "LOCATION"
                    mvarNewBill.Location = objRootElem.childNodes.Item(plID - 1).childNodes(j).Text
                Case "PAYTO"
                    mvarNewBill.PayTo = objRootElem.childNodes.Item(plID - 1).childNodes(j).Text
            End Select
        Next j
    End If
    mvarNewBill.InitForm btm_Edit
    mvarNewBill.Show vbModal
    Set mvarNewBill = Nothing
   
    If m_NewBillData.Mode = btm_Save Then
        If objDom.Load(Me.DataFile) Then
            Set objRootElem = objDom.documentElement
            For j = 0 To objRootElem.childNodes.Item(plID - 1).childNodes.length - 1
            Select Case UCase(objRootElem.childNodes.Item(plID - 1).childNodes(j).nodeName)
                Case "DUEDATE"
                    objRootElem.childNodes.Item(plID - 1).childNodes(j).Text = m_NewBillData.DueDate
                Case "AMOUNT"
                    objRootElem.childNodes.Item(plID - 1).childNodes(j).Text = m_NewBillData.Amount
                Case "LOCATION"
                    objRootElem.childNodes.Item(plID - 1).childNodes(j).Text = m_NewBillData.Location
                Case "PAYTO"
                    objRootElem.childNodes.Item(plID - 1).childNodes(j).Text = m_NewBillData.PayTo
            End Select
        Next j
            
            objDom.save Me.DataFile
            RaiseEvent RefreshGrid(Me.DataFile)
        End If
            
    End If
End Sub
Public Sub DeleteBill(plID As Long)
    Dim objDom As DOMDocument
    Dim objRootElem As IXMLDOMElement
    
    Dim objCurNode As IXMLDOMNode
    Dim objDelNode As IXMLDOMNode
    
    Dim sMsg As String
    Dim j As Long, i As Long
    
    Set objDom = New DOMDocument
    
    sMsg = "Delete bill from list?" & vbCrLf & vbCrLf
    If objDom.Load(Me.DataFile) Then
         Set objRootElem = objDom.documentElement
                    
            For j = 0 To objRootElem.childNodes.Item(plID - 1).childNodes.length - 1
                sMsg = sMsg & objRootElem.childNodes.Item(plID - 1).childNodes(j).nodeName & " = " & objRootElem.childNodes.Item(plID - 1).childNodes(j).Text & vbCrLf
            Next j
            
            
            If MsgBox(sMsg, vbQuestion + vbYesNo, "Delete Bill") = vbYes Then
                For j = 0 To objRootElem.childNodes.Item(plID - 1).childNodes.length - 1
                    'Set currNode = root.childNodes.item(1)
                    Set objCurNode = objRootElem.childNodes.Item(plID - 1)
                    'Set oldChild = currNode.removeChild(currNode.childNodes.Item(1))
                    If Not (objCurNode.childNodes(plID - 1) Is Nothing) Then
                        Set objDelNode = objCurNode.removeChild(objCurNode.childNodes(plID - 1))
                    End If
                    
                Next j
                Set objDelNode = objRootElem.removeChild(objRootElem.childNodes.Item(plID - 1))
                objDom.save Me.DataFile
                RaiseEvent RefreshGrid(Me.DataFile)
            End If
            
    End If
End Sub
Public Sub FillListView(pListView As ListView, Optional psXMLFile As String)

    Dim objDom As DOMDocument
    Dim objRootElem As IXMLDOMElement
    Dim objMemberElem As IXMLDOMElement
    Dim i As Long, j As Long
    Dim itmX As ListItem
    
    pListView.ListItems.Clear
    
    Set objDom = New DOMDocument
    
    If objDom.Load(psXMLFile) Then
         Set objRootElem = objDom.documentElement
    
        
        
        For i = 0 To objRootElem.childNodes.length - 1
            Set itmX = pListView.ListItems.Add(i + 1)
            itmX.Text = Trim$(str$(i + 1))
                    
            For j = 0 To objRootElem.childNodes.Item(i).childNodes.length - 1
                itmX.SubItems(j + 1) = objRootElem.childNodes.Item(i).childNodes(j).Text
            Next j
            
            Set itmX = Nothing
        Next i
    End If
        
End Sub


Public Sub AddBill()

   Dim objDom As DOMDocument
   Dim objRootElem As IXMLDOMElement
   Dim objMemberElem As IXMLDOMElement
   Dim objMemberRel As IXMLDOMAttribute
   Dim objMemberName As IXMLDOMElement
   
    
    Set mvarNewBill = New frmNewBill
    Load mvarNewBill
    mvarNewBill.InitForm btm_New
    mvarNewBill.Show vbModal
    Set mvarNewBill = Nothing
   
    
    If m_NewBillData.Mode = btm_None Then
        Exit Sub
    End If
    
   Set objDom = New DOMDocument
   
   objDom.Load Me.DataFile
   
    If Len(objDom.xml) = 0 Then
        ' Creates root element
        Set objRootElem = objDom.createElement("Bills")
        objDom.appendChild objRootElem
   
        ' Creates Member element
        Set objMemberElem = objDom.createElement("Current")
        objRootElem.appendChild objMemberElem
   
        Set objMemberName = objDom.createElement("DueDate")
        objMemberElem.appendChild objMemberName
        objMemberName.Text = m_NewBillData.DueDate
   
        Set objMemberName = objDom.createElement("PayTo")
        objMemberElem.appendChild objMemberName
        objMemberName.Text = m_NewBillData.PayTo
        
        Set objMemberName = objDom.createElement("Amount")
        objMemberElem.appendChild objMemberName
        objMemberName.Text = m_NewBillData.Amount
   
        Set objMemberName = objDom.createElement("Location")
        objMemberElem.appendChild objMemberName
        objMemberName.Text = m_NewBillData.Location
   
        
    Else
        'append
        Set objRootElem = objDom.documentElement
        'Set objMemberElem = objDom.firstChild
        
        ' Creates Member element
        Set objMemberElem = objDom.createElement("Current")
        objRootElem.appendChild objMemberElem
   
        Set objMemberName = objDom.createElement("DueDate")
        objMemberElem.appendChild objMemberName
        objMemberName.Text = m_NewBillData.DueDate
   
        Set objMemberName = objDom.createElement("PayTo")
        objMemberElem.appendChild objMemberName
        objMemberName.Text = m_NewBillData.PayTo
   
        Set objMemberName = objDom.createElement("Amount")
        objMemberElem.appendChild objMemberName
        objMemberName.Text = m_NewBillData.Amount
   
        Set objMemberName = objDom.createElement("Location")
        objMemberElem.appendChild objMemberName
        objMemberName.Text = m_NewBillData.Location
     
          
    End If
   objDom.save Me.DataFile
   Set objDom = Nothing
   
   RaiseEvent RefreshGrid(Me.DataFile)
   
End Sub


Public Property Let LogFile(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LogFile = 5
    mvarLogFile = vData
    mvarLogFile = Replace(mvarLogFile, "\\", "\")
    
    If Not FileExists(mvarLogFile) Then
        Dim FileHandler As Integer
        
        FileHandler = FreeFile
        Open mvarLogFile For Append As FileHandler
        Close FileHandler
    End If
End Property


Public Property Get LogFile() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LogFile
    LogFile = mvarLogFile
End Property



Public Property Let DataFile(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DataFile = 5
    mvarDataFile = vData
    mvarDataFile = Replace(mvarDataFile, "\\", "\")
    
    If Not FileExists(mvarDataFile) Then
        Dim FileHandler As Integer
        
        FileHandler = FreeFile
        Open mvarDataFile For Append As FileHandler
        Close FileHandler
    End If
    
End Property


Public Property Get DataFile() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DataFile
    DataFile = mvarDataFile
End Property




Private Sub mvarNewBill_SaveData(psPayTo As String, psDueDate As String, psAmount As String, psLocation As String, pMode As enmBillTrackerMode)
    m_NewBillData.Amount = psAmount
    m_NewBillData.DueDate = psDueDate
    m_NewBillData.Location = psLocation
    m_NewBillData.PayTo = psPayTo
    m_NewBillData.Mode = pMode
End Sub
