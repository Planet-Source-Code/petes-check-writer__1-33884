VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDragControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Unfortunately, a fair amount of additional logic
'is required only for line controls

#Const ADD_LINE_LOGIC = True

'Private Type PointAPI
'  X As Long
'  Y As Long
'End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function ClientToScreen Lib "user32" _
(ByVal hwnd As Long, lpPoint As PointAPI) As Long

Private Declare Function ScreenToClient Lib "user32" _
(ByVal hwnd As Long, lpPoint As PointAPI) As Long

Private m_Rect As RECT

#If ADD_LINE_LOGIC Then

Private Const SWAP_NONE = &H0
Private Const SWAP_X = &H1
Private Const SWAP_Y = &H2
Private m_fRectSwap As Integer

#End If

Public Property Let Left(NewLeft As Long)
m_Rect.Left = NewLeft
End Property

Public Property Get Left() As Long
Left = m_Rect.Left
End Property

Public Property Let Top(NewTop As Long)
m_Rect.Top = NewTop
End Property

Public Property Get Top() As Long
Top = m_Rect.Top
End Property

Public Property Let Right(NewRight As Long)
m_Rect.Right = NewRight
End Property

Public Property Get Right() As Long
Right = m_Rect.Right
End Property

Public Property Let Bottom(NewBottom As Long)
m_Rect.Bottom = NewBottom
End Property

Public Property Get Bottom() As Long
Bottom = m_Rect.Bottom
End Property

Public Property Let Width(NewWidth As Long)
m_Rect.Right = m_Rect.Left + NewWidth
End Property

Public Property Get Width() As Long
Width = m_Rect.Right - m_Rect.Left
End Property

Public Property Let Height(NewHeight As Long)
m_Rect.Bottom = m_Rect.Top + NewHeight
End Property

Public Property Get Height() As Long
Height = m_Rect.Bottom - m_Rect.Top
End Property

Public Sub SetRectToCtrl(ctl As Control)

#If ADD_LINE_LOGIC Then

'Reset swap flags
m_fRectSwap = SWAP_NONE
If TypeOf ctl Is Line Then
  m_Rect.Left = ctl.X1
  m_Rect.Top = ctl.Y1
  m_Rect.Right = ctl.X2
  m_Rect.Bottom = ctl.Y2
  'Need valid rect for hit testing but
  'must swap back in SetCtrlToRect
  If m_Rect.Left > m_Rect.Right Then
    m_fRectSwap = m_fRectSwap Or SWAP_X
  End If
    If m_Rect.Top > m_Rect.Bottom Then
      m_fRectSwap = m_fRectSwap Or SWAP_Y
    End If
    'Normalize if needed
    If m_fRectSwap <> SWAP_NONE Then
      NormalizeRect
    End If
  Else
    m_Rect.Left = ctl.Left
    m_Rect.Top = ctl.Top
    m_Rect.Right = ctl.Left + ctl.Width
    m_Rect.Bottom = ctl.Top + ctl.Height
  End If

#Else

m_Rect.Left = ctl.Left
m_Rect.Top = ctl.Top
m_Rect.Right = ctl.Left + ctl.Width
m_Rect.Bottom = ctl.Top + ctl.Height

#End If

End Sub


Public Sub SetCtrlToRect(ctl As Control)

#If ADD_LINE_LOGIC Then

If TypeOf ctl Is Line Then
'Restore normalized rectangle if needed
If m_fRectSwap And SWAP_X Then
  ctl.X1 = m_Rect.Right
  ctl.X2 = m_Rect.Left
Else
  ctl.X1 = m_Rect.Left
  ctl.X2 = m_Rect.Right
End If

If m_fRectSwap And SWAP_Y Then
  ctl.Y1 = m_Rect.Bottom
  ctl.Y2 = m_Rect.Top
Else
  ctl.Y1 = m_Rect.Top
  ctl.Y2 = m_Rect.Bottom
End If

'Force to valid rectangle
NormalizeRect

Else

'Force to valid rectangle

NormalizeRect

ctl.Move m_Rect.Left, m_Rect.Top, Width, Height

End If

#Else

'Force to valid rectangle

NormalizeRect
ctl.Move m_Rect.Left, m_Rect.Top, Width, Height

#End If

End Sub

Public Sub ScreenToTwips(ctl As Object)

Dim pt As PointAPI

pt.x = m_Rect.Left
pt.y = m_Rect.Top
ScreenToClient ctl.Parent.hwnd, pt
m_Rect.Left = pt.x * Screen.TwipsPerPixelX
m_Rect.Top = pt.y * Screen.TwipsPerPixelX
pt.x = m_Rect.Right
pt.y = m_Rect.Bottom
ScreenToClient ctl.Parent.hwnd, pt
m_Rect.Right = pt.x * Screen.TwipsPerPixelX
m_Rect.Bottom = pt.y * Screen.TwipsPerPixelX
End Sub

Public Sub TwipsToScreen(ctl As Object)
Dim pt As PointAPI

pt.x = m_Rect.Left / Screen.TwipsPerPixelX
pt.y = m_Rect.Top / Screen.TwipsPerPixelX
ClientToScreen ctl.Parent.hwnd, pt
m_Rect.Left = pt.x
m_Rect.Top = pt.y
pt.x = m_Rect.Right / Screen.TwipsPerPixelX
pt.y = m_Rect.Bottom / Screen.TwipsPerPixelX
ClientToScreen ctl.Parent.hwnd, pt
m_Rect.Right = pt.x
m_Rect.Bottom = pt.y
End Sub

Public Sub NormalizeRect()
Dim nTemp As Long

If m_Rect.Left > m_Rect.Right Then
nTemp = m_Rect.Right
m_Rect.Right = m_Rect.Left
m_Rect.Left = nTemp
End If
If m_Rect.Top > m_Rect.Bottom Then
nTemp = m_Rect.Bottom
m_Rect.Bottom = m_Rect.Top
m_Rect.Top = nTemp
End If
End Sub

Public Function PtInRect(x As Single, y As Single) As Integer
If x >= m_Rect.Left And x < m_Rect.Right And _
y >= m_Rect.Top And y < m_Rect.Bottom Then
PtInRect = True
Else
PtInRect = False
End If
End Function

