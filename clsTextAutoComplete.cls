VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextAutoComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private WithEvents m_txtComplete As TextBox
Attribute m_txtComplete.VB_VarHelpID = -1
Private m_strDelimeter As String
Private m_strList As String


Private Sub m_txtComplete_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Dim i As Integer
    Dim strSearchText As String
    Dim intDelimented As Integer
    Dim intLength As Integer
    Dim varArray As Variant
    


    With m_txtComplete
        If KeyCode <> vbKeyBack And KeyCode > 48 Then


        If InStr(1, m_strList, .Text, vbTextCompare) <> 0 Then
            varArray = Split(m_strList, m_strDelimeter)
            


            For i = 0 To UBound(varArray)
                strSearchText = Trim(varArray(i))
                
                If InStr(1, strSearchText, .Text, vbTextCompare) And (Left$(.Text, 1) = Left$(strSearchText, 1)) And .Text <> "" Then
                .SelText = ""
                .SelLength = 0
                intLength = Len(.Text)
                .Text = .Text & Right$(strSearchText, Len(strSearchText) - Len(.Text))
                .SelStart = intLength
                .SelLength = Len(.Text)
                Exit Sub
            End If
            
        Next i
    End If
End If
End With

End Sub


Public Property Get CompleteTextbox() As TextBox
    Set CompleteTextbox = m_txtComplete
End Property
Public Property Set CompleteTextbox(ByRef txt As TextBox)
Set m_txtComplete = txt
End Property


Public Property Get SearchList() As String
    SearchList = m_strList
End Property


Public Property Let SearchList(ByVal str As String)
    m_strList = str
End Property


Public Property Get Delimeter() As String
    Delimeter = m_strDelimeter
End Property


Public Property Let Delimeter(ByVal str As String)
    m_strDelimeter = str
End Property

