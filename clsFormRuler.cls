VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormRuler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Windows declarations

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type



Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd _
As Long, lpRect As RECT) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
ByVal hdc As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As _
Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, _
ByVal nDrawMode As Long) As Long

Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, _
ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Const NULL_BRUSH = 5
Private Const PS_SOLID = 0
Private Const R2_NOT = 6


'local variable(s) to hold property value(s)
Private mvarDragRect As clsDragControls
Private mvarMoveX As Single 'local copy
Private mvarMoveY As Single 'local copy
Private mvarX As Single 'local copy
Private mvarY As Single 'local copy
Private mvarResizeX As Single 'local copy
Private mvarResizeY As Single 'local copy
Private mvarIsMoving As Boolean 'local copy
Private mvarForm As frmCheckRuler 'local copy

Private Const RulerSize As Integer = 32
Private Const GridSize = 0.0834

Private mvarSaveX1 As Single 'local copy
Private mvarSaveY1 As Single 'local copy
'local variable(s) to hold property value(s)
Private mvarCurrCtl As Control 'local copy
Public Enum enmControlState
    StateNothing = 0
    StateDragging = 1
    StateSizing = 2
End Enum

Private mvarDragState As enmControlState 'local copy
Private mvarDragHandle As Integer 'local copy

Public Enum enmMode
    md_None = 0
    md_New = 1
    md_Edit = 2
    md_View = 3
End Enum

Private mvarMode As enmMode 'local copy
Private mvarOpenFile As String 'local copy
Private mvarDesignChanged As Boolean 'local copy
Private mvarDesignMode As Boolean 'local copy
Private mvarMDIForm As frmMain 'local copy
'local variable(s) to hold property value(s)
'local variable(s) to hold property value(s)
Private mvarDragPoint As clsPoint 'local copy
Public Sub PrintDesign()

On Error GoTo PrintDesign_Error

    ShowMessage anm_Printing, "Printing Check Design: " & Me.OpenFile, "Printing Design Sample"
    
    
    If Not SelectPrinter() Then
        Unload frmShowAnim
        Exit Sub
    End If
    
    
    'print check outline
    SetPrintCheckRect
    SetPrintTestFields
    SetPrintRuler
    SetPrintInfo
    
    Printer.EndDoc
    Unload frmShowAnim

Exit Sub
PrintDesign_Error:
    Unload frmShowAnim
    MsgBox "Error #: " & str(Err) & vbCrLf & "Error Desc: " & Error & vbCrLf & "Printing Design Sample", vbExclamation

End Sub
Private Sub SetPrintInfo()

      With Printer
        .ScaleMode = vbInches
        .CurrentX = 0
        .CurrentY = 3.6
        
        .FontName = "Arial"
        .FontSize = 16
        .FontBold = True
        .FontUnderline = True
        Printer.Print "Design Sample: " & Me.OpenFile & vbCrLf & vbCrLf
        .FontBold = False
        .FontUnderline = False
        
        .FontName = "Courier New"
        .FontSize = 8
        .FontBold = True
        .FontUnderline = True
        Printer.Print "Instructions"
        .FontBold = False
        .FontUnderline = False
        Printer.Print "1. Cut the dotted lines.  Do not start before the solid line.  Start ** ON ** the solid line."
        Printer.Print "    Do not cut past the solid line.  Stop ** ON ** the solid line."
        Printer.Print ""
        Printer.Print "2. Place the four corners of the check into the slits in the check holder.  Hold up to the light."
        Printer.Print "    Verify position of fields.  Adjust in design file accordingly!"
        
        Printer.Print vbCrLf
                
        Printer.Font = "Arial"
        Printer.FontSize = 8
        
        .FontBold = True
        .FontUnderline = True
        Printer.Print "Standard Software Disclaimer" & vbCrLf & vbCrLf
        .FontUnderline = False
        .FontBold = False
        Printer.Print "While Pete Sral makes every effort to deliver high quality products, we do not guarantee"
        Printer.Print "that our products are free from defects.  Our software is provided 'as is,' and you use"
        Printer.Print "the software at your own risk. We make no warranties as to performance, merchantability, "
        Printer.Print "fitness for a particular purpose, or any other warranties whether expressed or implied.  "
        Printer.Print "No oral or written communication from or information provided by Pete Sral shall create a "
        Printer.Print "warranty. Under no circumstances shall Pete Sral be liable for direct, indirect, special, "
        Printer.Print "incidental, or consequential damages resulting from the use, misuse, or inability to use "
        Printer.Print "this software, even if Pete Sral has been advised of the possibility of such damages."
        Printer.Print vbCrLf & vbCrLf

        Printer.Print Chr(169) & Format(Now, "yyyy") & " Pete Sral, All rights reserved"
        Printer.Print "Check Writer: Version " & App.Major & "." & App.Minor & "." & App.Revision
        Printer.Print ""
        Printer.Print "Thank you for using the Check Writer software! - Have a nice day!"
        
        .CurrentX = 0
        .CurrentY = 0
        .FontName = "Courier New"
        .ScaleMode = vbTwips
    End With
        
End Sub
Private Sub SetPrintRuler()

    'Printer.PaintPicture Picture1.Picture, .CurrentX, .CurrentY
    
    With Printer
        .ScaleMode = vbInches
        .CurrentX = 0
        .CurrentY = 2.85
        Printer.PaintPicture LoadResPicture("HRuler", vbResBitmap), .CurrentX, .CurrentY
        
        .CurrentX = 6.05
        .CurrentY = 0
        Printer.PaintPicture LoadResPicture("VRuler", vbResBitmap), .CurrentX, .CurrentY
    End With
        
End Sub
Private Sub SetPrintTestFields()

Dim i As Integer

    With Printer
        .ScaleMode = vbTwips
        ''.ColorMode = 1
        .Font.Name = "Courier New"
        .Font.Size = 10
        
        For i = 0 To 6
            .CurrentX = 0
            .CurrentY = 0
            
            .CurrentX = mvarForm.lblPosition(i).Left - piRulerOffsetX
            .CurrentY = mvarForm.lblPosition(i).Top - piRulerOffsetY
            Printer.Print Trim$(mvarForm.lblPosition(i).Caption)
        Next i
    End With
End Sub
Private Sub SetPrintCheckRect()

With Printer
    
        .DrawStyle = vbSolid
        .ScaleMode = vbInches
        ''.ColorMode = 1
        .CurrentX = 0
        .CurrentY = 0
        
        'Printer.Print "X"
        'top
        Printer.Line (.CurrentX, .CurrentY)-(.CurrentX + 6, .CurrentY)
               
        'right
        Printer.Line (.CurrentX, .CurrentY)-(.CurrentX, .CurrentY + 2.8)
        
        .CurrentX = 0
        .CurrentY = 0
        
        'left
        Printer.Line (.CurrentX, .CurrentY)-(.CurrentX, .CurrentY + 2.8)
        
        'bottom
        Printer.Line (.CurrentX, .CurrentY)-(.CurrentX + 6, .CurrentY)
        
        
        .CurrentX = 0
        .CurrentY = 0
        
        .DrawStyle = vbDot
         
         
        'top left to left
        Printer.Line (.CurrentX + 0.5, .CurrentY)-(.CurrentX, .CurrentY + 0.5)
        
        .CurrentX = 6
        .CurrentY = 0
        
        'top right to right
        Printer.Line (.CurrentX - 0.5, .CurrentY)-(.CurrentX, .CurrentY + 0.5)
        
        .CurrentX = 0
        .CurrentY = 2.8
        
        
        'bottom left to left
        Printer.Line (.CurrentX, .CurrentY - 0.5)-(.CurrentX + 0.5, .CurrentY)
        
        .CurrentX = 6
        .CurrentY = 2.8
        
        'top right to right
        Printer.Line (.CurrentX, .CurrentY - 0.5)-(.CurrentX - 0.5, .CurrentY)
        
End With
End Sub
Public Property Set DragPoint(ByVal vData As clsPoint)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DragPoint = Form1
    Set mvarDragPoint = vData
End Property


Public Property Get DragPoint() As clsPoint
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DragPoint
    Set DragPoint = mvarDragPoint
End Property

Public Sub Form_MouseMove(x As Single, y As Single)
Dim nWidth As Single, nHeight As Single
Dim pt As PointAPI

    If Me.DragState = StateDragging Then
        'Save dimensions before modifying rectangle
        nWidth = Me.DragRect.Right - Me.DragRect.Left
        nHeight = Me.DragRect.Bottom - Me.DragRect.Top

        'Get current mouse position in screen coordinates
        GetCursorPos pt

        'Hide existing rectangle
        Me.DrawDragRect

        'Update drag rectangle coordinates
    
        Me.DragRect.Left = pt.x - Me.DragPoint.x
        Me.DragRect.Top = pt.y - Me.DragPoint.y
        Me.DragRect.Right = Me.DragRect.Left + nWidth
        Me.DragRect.Bottom = Me.DragRect.Top + nHeight

        'Draw new rectangle
        Me.DrawDragRect
    ElseIf Me.DragState = StateSizing Then
        'Get current mouse position in screen coordinates
        GetCursorPos pt

        'Hide existing rectangle
        Me.DrawDragRect

        'Action depends on handle being dragged
        Select Case Me.DragHandle
            Case 0
                Me.DragRect.Left = pt.x
                Me.DragRect.Top = pt.y
            
            Case 1
                Me.DragRect.Top = pt.y
            
            Case 2
                Me.DragRect.Right = pt.x
                Me.DragRect.Top = pt.y

            Case 3
                Me.DragRect.Right = pt.x

            Case 4
                Me.DragRect.Right = pt.x
                Me.DragRect.Bottom = pt.y

            Case 5
                Me.DragRect.Bottom = pt.y

            Case 6
                Me.DragRect.Left = pt.x
                Me.DragRect.Bottom = pt.y
            
            Case 7
                Me.DragRect.Left = pt.x

        End Select

        'Draw new rectangle
        Me.DrawDragRect
    End If

    Me.x = x
    Me.y = y
    
   
    
    If Not Me.IsMoving Then
        Me.HandlePositionIndicator x, y
    End If
    
    If Len(mvarForm.sbStatusBar.Panels(2).Text) <> 0 Then
        mvarForm.sbStatusBar.Panels(2).Text = ""
    End If
    
    If Len(mvarForm.sbStatusBar.Panels(1).Text) <> 0 Then
        mvarForm.sbStatusBar.Panels(1).Text = ""
    End If
    
   'StatusBar1.Panels(2).Text = "X = " & TwipToInch(X - piRulerOffsetX) & " Y = " & TwipToInch(Y - piRulerOffsetY)
End Sub
Public Sub Form_MouseUp(Button As Integer)

If Button = vbLeftButton Then

        If Not (Me.CurrCtl Is Nothing) Then
            If Me.CurrCtl.Name = "Shape1" Or Me.CurrCtl.Name = "Shape2" Then
                Exit Sub
            End If
        End If
        
        If Me.DragState = StateDragging Or Me.DragState = StateSizing Then
            'Hide drag rectangle
            Me.DrawDragRect

            'Move control to new location
            Me.DragRect.ScreenToTwips Me.CurrCtl
            Me.DragRect.SetCtrlToRect Me.CurrCtl

            'Restore sizing handles
            Me.ShowHandles True

            'Free mouse movement
            ClipCursor ByVal 0&

            'Release mouse capture
            ReleaseCapture

            'Reset drag state
            Me.DragState = StateNothing
        End If
    End If
End Sub
Public Sub Form_MouseDown(Button As Integer)

On Error Resume Next

Dim i As Integer

    If Button = vbLeftButton And Me.DesignMode Then
        'Hit test over light-weight (non-windowed) mvarform.Controls
            For i = 0 To (mvarForm.Controls.Count - 1)
                'Check for visible, non-menu mvarform.Controls
                '[Note 1]
                'If any of the sizing handle mvarform.Controls are under the mouse
                'pointer, then they must not be visible or else they would
                'have already intercepted the MouseDown event.

                '[Note 2]
                'This code will fail if you have a control such as the
                'Timer control which has no Visible property. You will
                'either need to make sure your form has no such mvarform.Controls
                'or add code to handle them.
                
                If (Not TypeOf mvarForm.Controls(i) Is Menu And Not TypeOf mvarForm.Controls(i) Is Shape) And mvarForm.Controls(i).Visible Then
                    Me.DragRect.SetRectToCtrl mvarForm.Controls(i)
                    If Me.DragRect.PtInRect(x, y) Then
                        DragBegin mvarForm.Controls(i)
                        Exit Sub
                    End If
                End If
               
        Next i

        'No control is active
        Set Me.CurrCtl = Nothing

        'Hide sizing handles
        Me.ShowHandles False
    End If
    
    

End Sub
Public Sub DrawDragRect()


'Draw drag rectangle. The API is used for efficiency and also
'because drag rectangle must be drawn on the screen DC in
'order to appear on top of all controls

Dim hPen As Long, hOldPen As Long
Dim hBrush As Long, hOldBrush As Long
Dim hScreenDC As Long, nDrawMode As Long

    'Get DC of entire screen in order to
    'draw on top of all controls
    hScreenDC = GetDC(0)

    'Select GDI object
    hPen = CreatePen(PS_SOLID, 2, 0)
    hOldPen = SelectObject(hScreenDC, hPen)
    hBrush = GetStockObject(NULL_BRUSH)
    hOldBrush = SelectObject(hScreenDC, hBrush)
    nDrawMode = SetROP2(hScreenDC, R2_NOT)

    'Draw rectangle
    Rectangle hScreenDC, Me.DragRect.Left, Me.DragRect.Top, _
    Me.DragRect.Right, Me.DragRect.Bottom

    'Restore DC
    SetROP2 hScreenDC, nDrawMode
    SelectObject hScreenDC, hOldBrush
    SelectObject hScreenDC, hOldPen
    ReleaseDC 0, hScreenDC

    'Delete GDI objects
    DeleteObject hPen
End Sub

Public Sub DragBegin(ctl As Control)
Dim rc As RECT
Dim pt As PointAPI

    'Hide any visible handles
    Me.ShowHandles False

    'Save reference to control being dragged
    Set Me.CurrCtl = ctl

    'Store initial mouse position
    GetCursorPos pt

    Me.DragPoint.x = pt.x
    Me.DragPoint.y = pt.y
    'Save control position (in screen coordinates)
    'Note: control might not have a window handle
    Me.DragRect.SetRectToCtrl Me.CurrCtl
    Me.DragRect.TwipsToScreen Me.CurrCtl

    'Make initial mouse position relative to control
    Me.DragPoint.x = Me.DragPoint.x - Me.DragRect.Left
    Me.DragPoint.y = Me.DragPoint.y - Me.DragRect.Top

    'Force redraw of form without sizing handles
    'before drawing dragging rectangle
    mvarForm.Refresh

    'Show dragging rectangle
    DrawDragRect

    'Indicate dragging under way
    Me.DragState = StateDragging

    'In order to detect mouse movement over any part of the form,
    'we set the mouse capture to the form and will process mouse
    'movement from the applicable form events
    ReleaseCapture 'This appears needed before calling SetCapture
    SetCapture mvarForm.hwnd

    'Limit cursor movement within form
    GetWindowRect mvarForm.hwnd, rc
    ClipCursor rc

End Sub




Public Sub ShowHandles(Optional bShowHandles As Boolean = True, Optional bUpdateTextFields As Boolean = True)

Dim i As Integer
Dim xFudge As Long, yFudge As Long
Dim nWidth As Long, nHeight As Long



    If bShowHandles And Not Me.CurrCtl Is Nothing Then

        With Me.DragRect
            'Save some calculations in variables for speed
            nWidth = (mvarForm.picHandle(0).Width \ 2)
            nHeight = (mvarForm.picHandle(0).Height \ 2)

            xFudge = (0.5 * Screen.TwipsPerPixelX)
            yFudge = (0.5 * Screen.TwipsPerPixelY)

            'Top Left
            mvarForm.picHandle(0).Move (.Left - nWidth) + xFudge, (.Top - nHeight) + yFudge

            'Bottom right
            mvarForm.picHandle(4).Move (.Left + .Width) - nWidth - xFudge, .Top + .Height - nHeight - yFudge

            'Top center
            mvarForm.picHandle(1).Move .Left + (.Width / 2) - nWidth, .Top - nHeight + yFudge

            'Bottom center
            mvarForm.picHandle(5).Move .Left + (.Width / 2) - nWidth, .Top + .Height - nHeight - yFudge

            'Top right
            mvarForm.picHandle(2).Move .Left + .Width - nWidth - xFudge, .Top - nHeight + yFudge
            
            'Bottom left
            mvarForm.picHandle(6).Move .Left - nWidth + xFudge, .Top + .Height - nHeight - yFudge

            'Center right
            mvarForm.picHandle(3).Move .Left + .Width - nWidth - xFudge, .Top + (.Height / 2) - nHeight

            'Center left
            mvarForm.picHandle(7).Move .Left - nWidth + xFudge, .Top + (.Height / 2) - nHeight

        End With
        
        
    End If

        If Not (Me.CurrCtl Is Nothing) And bUpdateTextFields Then
            mvarForm.lblFieldName.Caption = "Field " & Me.CurrCtl.Tag & " Coordinates"
            mvarForm.txtXLeft.Text = Format((Me.CurrCtl.Left - piRulerOffsetX) / 1440, "#.##")
            mvarForm.txtYTop.Text = Format((Me.CurrCtl.Top - piRulerOffsetY) / 1440, "#.##")
            mvarForm.txtWidth.Text = Format((Me.CurrCtl.Width) / 1440, "#.##")
        End If
        
    'Show or hide each handle

    For i = 0 To 7
        mvarForm.picHandle(i).Visible = bShowHandles
    Next i

    mvarForm.Refresh
    
End Sub
Public Sub DragInit()

Dim i As Integer, xHandle As Single, yHandle As Single

    'Use black Picture box controls for 8 sizing handles
    'Calculate size of each handle
    xHandle = 5 * Screen.TwipsPerPixelX
    yHandle = 5 * Screen.TwipsPerPixelY

    'Load array of handles until we have 8
    For i = 0 To 7
        'If i <> 0 Then
            'Load mvarform.picHandle(i)
        'End If
        mvarForm.picHandle(i).Width = xHandle
        mvarForm.picHandle(i).Height = yHandle

        'Must be in front of other controls
        mvarForm.picHandle(i).ZOrder

    Next i
    
    For i = 0 To 6
        mvarForm.lblPosition(i).Enabled = False
    Next i

    'Set mousepointers for each sizing handle
    mvarForm.picHandle(0).MousePointer = vbSizeNWSE
    mvarForm.picHandle(1).MousePointer = vbSizeNS
    mvarForm.picHandle(2).MousePointer = vbSizeNESW
    mvarForm.picHandle(3).MousePointer = vbSizeWE
    mvarForm.picHandle(4).MousePointer = vbSizeNWSE
    mvarForm.picHandle(5).MousePointer = vbSizeNS
    mvarForm.picHandle(6).MousePointer = vbSizeNESW
    mvarForm.picHandle(7).MousePointer = vbSizeWE

    'Initialize current control
    Set Me.CurrCtl = Nothing
    
End Sub
Public Sub DragEnd()
Dim i As Integer

     For i = 0 To 6
        mvarForm.lblPosition(i).Enabled = False
     Next i
        
        
    'Clears any current drag mode and hides sizing handles
    Set Me.CurrCtl = Nothing
    Me.ShowHandles False
    Me.DragState = StateNothing
End Sub

Public Property Set MDIForm(ByVal vData As frmMain)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MDIForm = Form1
    Set mvarMDIForm = vData
End Property


Public Property Get MDIForm() As frmMain
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MDIForm
    Set MDIForm = mvarMDIForm
End Property



Public Property Let DesignMode(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DesignMode = 5

Dim bSave As Boolean
Dim i As Integer

    bSave = Me.DesignChanged
    Me.DesignChanged = False
    mvarDesignMode = vData
    
    If mvarDesignMode Then
        mvarMode = md_None
        Me.Mode = md_Edit
    Else
        Me.DragEnd
        mvarMode = md_View
        Me.Mode = md_View
        mvarForm.lblFieldName.Caption = "Field Coordinates"
    End If
    
    Me.DesignChanged = bSave
    mvarForm.mnuDesignDesign.Checked = mvarDesignMode
    mvarMDIForm.tbDesign.Buttons("Edit").Value = IIf(mvarDesignMode = True, tbrPressed, tbrUnpressed)
    
    If Me.DesignChanged Then
        With mvarForm
            .mnuDesignSave.Enabled = mvarDesignChanged
            .mnuDesignSaveAs.Enabled = mvarDesignChanged
            mvarMDIForm.tbDesign.Buttons("Save").Enabled = .mnuDesignSave.Enabled
        End With
    End If
       
     For i = 0 To 6
        mvarForm.lblPosition(i).Enabled = mvarDesignMode
    Next i
    
End Property


Public Property Get DesignMode() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DesignMode
    DesignMode = mvarDesignMode
End Property



Public Property Let DesignChanged(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DesignChanged = 5
    mvarDesignChanged = vData
End Property


Public Property Get DesignChanged() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DesignChanged
    DesignChanged = mvarDesignChanged
End Property



Public Property Let OpenFile(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OpenFile = 5
    If InStr(1, vData, "\") = 0 Then
        mvarOpenFile = App.path & "\" & vData
        mvarOpenFile = Replace(mvarOpenFile, "\\", "\")
    End If
    
    mvarOpenFile = vData
End Property


Public Property Get OpenFile() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OpenFile
    OpenFile = mvarOpenFile
End Property
Public Sub SaveDesign()
    DoSave Me.OpenFile
End Sub
Public Sub SaveDesignAs()

Dim cSaveFile As New clsCmDlg
    
    cSaveFile.DefaultExt = "ckc"
    cSaveFile.DialogTitle = "Save check coordinates"
    cSaveFile.FileName = ""
    cSaveFile.Filter = "Check Coordinates (*ckc)|*.ckc"
    cSaveFile.FilterIndex = 1
    cSaveFile.InitDir = App.path
    cSaveFile.ShowSave
    
    If Len(cSaveFile.FileName) <> 0 Then
        If FileExists(cSaveFile.FileName) Then
            If MsgBox("Check design file: " & cSaveFile.FileName & " already exists!" & vbCrLf & "Do you want to overwrite this file?", vbQuestion + vbYesNo, "File Exists") = vbNo Then
                Exit Sub
            End If
        End If
        
        DoSave cSaveFile.FileName
    End If
    
    Set cSaveFile = Nothing
    
End Sub
Private Sub DoSave(Optional psFileSaveAs As String)
    Dim cSave As New clsIniFile
    Dim i As Integer
    Dim sFileName As String
        
        If Len(psFileSaveAs) <> 0 Then
            sFileName = psFileSaveAs
        Else
            sFileName = Me.OpenFile
        End If
        
        
        If Len(sFileName) <> 0 And sFileName <> "New" Then
            
            With cSave
                If Right(sFileName, 4) <> ".ckc" Then
                    sFileName = sFileName & ".ckc"
                End If
            
                
                .FullPath = sFileName
            
                .SectionName = "Version"
                .PutString "Major", App.Major
                .PutString "Minor", App.Minor
                .PutString "Revision", App.Revision
            
                .SectionName = "Settings"
                .PutString "Path", App.path
            
                For i = 0 To 6
                    .SectionName = mvarForm.lblPosition(i).Tag
                    .PutString "Left", mvarForm.lblPosition(i).Left
                    .PutString "Top", mvarForm.lblPosition(i).Top
                    .PutString "Width", mvarForm.lblPosition(i).Width
                    .PutString "Height", mvarForm.lblPosition(i).Height
                Next i
        
                mvarForm.Caption = "Check Design: " & .ExtractName(.FileName, exnFullName)
                Me.OpenFile = .FullPath
            End With
        Else
            'MsgBox "Please specify a file name to save!", vbExclamation
            Me.SaveDesignAs
        End If
        
       
       
        Set cSave = Nothing
        Me.DesignChanged = False
        mvarMode = md_View
               
        Me.DesignMode = False
        
       
        
End Sub
Public Sub OpenDesign(Optional psFullPath As String)

   Dim cOpen As New clsIniFile
    Dim cOpenFile As New clsCmDlg
    Dim i As Integer
    Dim sFilePath As String
      
      Me.DesignMode = False
      Me.Mode = md_None
        
        sFilePath = psFullPath
        
        If Len(sFilePath) = 0 Then
            cOpenFile.DefaultExt = "ckc"
            cOpenFile.DialogTitle = "Open check file"
            cOpenFile.FileName = ""
            cOpenFile.Filter = "Check Coordinates (*ckc)|*.ckc"
            cOpenFile.FilterIndex = 1
            cOpenFile.InitDir = App.path
            cOpenFile.ShowOpen
        
            sFilePath = cOpenFile.FileName
        End If
        
        If Len(sFilePath) <> 0 Then
            With cOpen
            
                If Right(sFilePath, 4) <> ".ckc" Then
                    sFilePath = sFilePath & ".ckc"
                End If
            
                .FullPath = sFilePath
            
                For i = 0 To 6
                    .SectionName = mvarForm.lblPosition(i).Tag
                    mvarForm.lblPosition(i).Left = .GetString("Left", (100) + piRulerOffsetX)
                    mvarForm.lblPosition(i).Top = .GetString("Top", (i * 100) + piRulerOffsetY)
                    mvarForm.lblPosition(i).Height = .GetString("Height", 300)
                    mvarForm.lblPosition(i).Width = .GetString("Width", 500)
                Next i
        
                mvarForm.Caption = "Check Design: " & .ExtractName(.FileName, exnFullName)
                Me.OpenFile = .FullPath
            End With
        
            If p_clsOptions.EditOnOpen Then
                mvarForm.mnuDesignDesign_Click
            Else
                Me.Mode = md_View
            End If
            
            Me.DesignChanged = False
        End If
        
       
        Set cOpenFile = Nothing
        Set cOpen = Nothing
        
      
        
        
End Sub

Public Sub NewDesign()

Dim cINi As New clsIniFile
Dim i As Integer
       
    
    
    Me.DesignMode = False
    Me.Mode = md_New
    
    With cINi
            
        .FullPath = App.path & "\new.chk"
        
        For i = 0 To 6
            .SectionName = mvarForm.lblPosition(i).Tag
             mvarForm.lblPosition(i).Left = .GetString("Left", (100) + piRulerOffsetX)
             mvarForm.lblPosition(i).Top = .GetString("Top", (i * 100) + piRulerOffsetY)
             mvarForm.lblPosition(i).Height = .GetString("Height", 300)
             mvarForm.lblPosition(i).Width = .GetString("Width", 500)
        Next i
        
        mvarForm.Caption = "Check Design: New"
        Me.OpenFile = "New"
        mvarForm.Refresh
        
        If p_clsOptions.EditOnOpen Then
                Me.Mode = md_Edit
                mvarForm.mnuDesignDesign_Click
        End If
        
        Me.DesignChanged = False
    End With
    
End Sub


Public Property Let Mode(ByVal vData As enmMode)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Mode = Form1
Dim i As Integer
    
    If (mvarMode = md_Edit Or mvarMode = md_New) And Me.DesignChanged Then
        If MsgBox("Save current check design?", vbQuestion + vbYesNo) = vbYes Then
            Me.SaveDesign
        End If
    End If
    
    Me.DesignChanged = False
    mvarMode = vData
    
    SetMDIStatusBarMode mvarMode
    
    For i = 0 To 6
        mvarForm.lblPosition(i).Visible = (mvarMode <> md_None)
        mvarForm.lblPosition(i).Enabled = (mvarMode = md_Edit)
    Next i
        
    Select Case mvarMode
         Case md_None '= 0
            With mvarForm
                .mnuDesignNew.Enabled = True
                .mnuDesignDesign.Enabled = False
                .mnuDesignOpen.Enabled = True
                .mnuDesignSave.Enabled = False
                .mnuDesignSaveAs.Enabled = False
                .mnuDesignPrintDesignSample.Enabled = False
            End With
                
         Case md_New '= 1
            With mvarForm
                .mnuDesignNew.Enabled = False
                .mnuDesignDesign.Enabled = True
                .mnuDesignOpen.Enabled = True
                .mnuDesignSave.Enabled = False Or mvarDesignChanged
                .mnuDesignSaveAs.Enabled = True
                .mnuDesignPrintDesignSample.Enabled = True
            End With
            
            
         Case md_Edit '= 2
             With mvarForm
                .mnuDesignNew.Enabled = True
                .mnuDesignDesign.Enabled = True
                .mnuDesignOpen.Enabled = True
                .mnuDesignSave.Enabled = Len(Me.OpenFile) <> 0
                .mnuDesignSaveAs.Enabled = True
                .mnuDesignPrintDesignSample.Enabled = True
            End With
         Case md_View '= 3
             With mvarForm
                .mnuDesignNew.Enabled = True
                .mnuDesignDesign.Enabled = True
                .mnuDesignOpen.Enabled = True
                .mnuDesignSave.Enabled = False Or mvarDesignChanged
                .mnuDesignSaveAs.Enabled = False Or mvarDesignChanged
                .mnuDesignPrintDesignSample.Enabled = True
            End With
    End Select
    
        
    With mvarForm
        mvarMDIForm.tbDesign.Buttons("New").Enabled = .mnuDesignNew.Enabled
        mvarMDIForm.tbDesign.Buttons("Edit").Enabled = .mnuDesignDesign.Enabled
        mvarMDIForm.tbDesign.Buttons("Open").Enabled = .mnuDesignOpen.Enabled
        mvarMDIForm.tbDesign.Buttons("Save").Enabled = .mnuDesignSave.Enabled
        mvarMDIForm.tbDesign.Buttons("Print").Enabled = .mnuDesignPrintDesignSample.Enabled
    End With
    
End Property


Public Property Get Mode() As enmMode
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Mode
    Mode = mvarMode
End Property




Public Property Set DragRect(ByVal vData As clsDragControls)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DragRect = Form1
    Set mvarDragRect = vData
End Property


Public Property Get DragRect() As clsDragControls
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DragRect
    Set DragRect = mvarDragRect
End Property



Public Property Let DragHandle(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DragHandle = 5
    mvarDragHandle = vData
End Property


Public Property Get DragHandle() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DragHandle
    DragHandle = mvarDragHandle
End Property



Public Property Let DragState(ByVal vData As enmControlState)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DragState = Form1
    mvarDragState = vData
End Property


Public Property Get DragState() As enmControlState
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DragState
    DragState = mvarDragState
End Property




Public Property Set CurrCtl(ByVal vData As Control)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.CurrCtl = Form1
    Set mvarCurrCtl = vData
End Property


Public Property Get CurrCtl() As Control
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CurrCtl
    Set CurrCtl = mvarCurrCtl
End Property



Public Property Let SaveY1(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SaveY1 = 5
    mvarSaveY1 = vData
End Property


Public Property Get SaveY1() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SaveY1
    SaveY1 = mvarSaveY1
End Property



Public Property Let SaveX1(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SaveX1 = 5
    mvarSaveX1 = vData
End Property


Public Property Get SaveX1() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SaveX1
    SaveX1 = mvarSaveX1
End Property




Public Sub Initialize(Optional pbRefresh As Boolean = False)

Dim i As Integer
Dim cINi As New clsIniFile

    mvarForm.HRuler.Width = mvarForm.ScaleWidth - mvarForm.HRuler.Left
    mvarForm.VRuler.Height = mvarForm.ScaleHeight - mvarForm.VRuler.Top
    
    
    
    Me.DrawHorizontalRuler
    Me.DrawVerticalRuler
    
    
    If Not pbRefresh Then
        'Me.DrawGrid
        'save defaults
        mvarForm.Width = 9540
        mvarForm.Height = 6375
    
        With cINi
            .FullPath = App.path & "\new.chk"
        
            .SectionName = "Version"
            .PutString "Major", App.Major
            .PutString "Minor", App.Minor
            .PutString "Revision", App.Revision
            
            .SectionName = "Settings"
            .PutString "Path", App.path
            
            For i = 0 To 6
                .SectionName = mvarForm.lblPosition(i).Tag
                .PutString "Left", mvarForm.lblPosition(i).Left
                .PutString "Top", mvarForm.lblPosition(i).Top
                .PutString "Width", mvarForm.lblPosition(i).Width
                .PutString "Height", mvarForm.lblPosition(i).Height
            Next i
        End With
        
        Set cINi = Nothing
        
        If Len(Trim$(p_clsOptions.DefaultCheckPath)) <> 0 Then
            Me.OpenDesign p_clsOptions.DefaultCheckPath
        Else
            Me.Mode = md_None
        End If
    End If
    
    
End Sub
Public Property Set Form(ByVal vData As frmCheckRuler)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Form = mvarform
    Set mvarForm = vData
End Property


'Public Property Get Form() As frmCheckRuler
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Form
'    Set Form = mvarForm
'End Property



Public Sub DrawHorizontalRuler()
    Dim Sincr As Single, RScale As Integer, i As Integer
    
    'Scalemode is in TWIPS 1440 per inch
    RScale = 1440
    
    'Number of segment across form
    Sincr = RScale / RulerSize
    Do While Sincr < mvarForm.HRuler.ScaleWidth
        'Number of sections
        For i = 1 To RulerSize
            'Size of Tics
            If i = RulerSize Then
                mvarForm.HRuler.Line (Sincr, 0)-(Sincr, mvarForm.HRuler.ScaleHeight)
                mvarForm.HRuler.CurrentY = 0
                mvarForm.HRuler.Print Int(Sincr / RScale)
            Else
                If i = Int(RulerSize * 0.75) Then
                    mvarForm.HRuler.Line (Sincr, mvarForm.HRuler.ScaleHeight - (mvarForm.HRuler.ScaleHeight * 0.25))-(Sincr, mvarForm.HRuler.ScaleHeight)
                Else
                    If i = Int(RulerSize * 0.5) Then
                        mvarForm.HRuler.Line (Sincr, mvarForm.HRuler.ScaleHeight - (mvarForm.HRuler.ScaleHeight * 0.5))-(Sincr, mvarForm.HRuler.ScaleHeight)
                    Else
                        If i = RulerSize * 0.25 Then
                            mvarForm.HRuler.Line (Sincr, mvarForm.HRuler.ScaleHeight - (mvarForm.HRuler.ScaleHeight * 0.25))-(Sincr, mvarForm.HRuler.ScaleHeight)
                        Else
                            mvarForm.HRuler.Line (Sincr, mvarForm.HRuler.ScaleHeight - (mvarForm.HRuler.ScaleHeight * 0.125))-(Sincr, mvarForm.HRuler.ScaleHeight)
                        End If
                    End If
                End If
            End If
            Sincr = Sincr + (RScale / RulerSize)
        Next
    Loop
End Sub

Public Sub DrawVerticalRuler()
    Dim Sincr As Single, RScale As Integer, i As Integer
    
    'Scalemode is in TWIPS 1440 per inch
    RScale = 1440
    
    'Number of segment across form
    Sincr = RScale / RulerSize
    
    Do While Sincr < mvarForm.VRuler.ScaleHeight
        'Number of sections
        For i = 1 To RulerSize
            'Size of Tics
            If i = RulerSize Then
                mvarForm.VRuler.Line (0, Sincr)-(mvarForm.VRuler.ScaleWidth, Sincr)
                mvarForm.VRuler.CurrentX = 0
                mvarForm.VRuler.Print Int(Sincr / RScale)
            Else
                If i = Int(RulerSize * 0.75) Then
                   ' mvarform.VRuler.Line (Sincr, mvarform.VRuler.ScaleWidth - (mvarform.VRuler.ScaleWidth * 0.25))-(Sincr, mvarform.VRuler.ScaleWidth)
                   mvarForm.VRuler.Line (mvarForm.VRuler.ScaleWidth - (mvarForm.VRuler.ScaleWidth * 0.25), Sincr)-(mvarForm.VRuler.ScaleWidth, Sincr)
                   
                Else
                    If i = Int(RulerSize * 0.5) Then
                   '     mvarform.VRuler.Line (Sincr, mvarform.VRuler.ScaleWidth - (mvarform.VRuler.ScaleWidth * 0.5))-(Sincr, mvarform.VRuler.ScaleWidth)
                      mvarForm.VRuler.Line (mvarForm.VRuler.ScaleWidth - (mvarForm.VRuler.ScaleWidth * 0.5), Sincr)-(mvarForm.VRuler.ScaleWidth, Sincr)
                    Else
                        If i = RulerSize * 0.25 Then
                    '        mvarform.VRuler.Line (Sincr, mvarform.VRuler.ScaleWidth - (mvarform.VRuler.ScaleWidth * 0.25))-(Sincr, mvarform.VRuler.ScaleWidth)
                    mvarForm.VRuler.Line (mvarForm.VRuler.ScaleWidth - (mvarForm.VRuler.ScaleWidth * 0.25), Sincr)-(mvarForm.VRuler.ScaleWidth, Sincr)
                        Else
                    '        mvarform.VRuler.Line (Sincr, mvarform.VRuler.ScaleWidth - (mvarform.VRuler.ScaleWidth * 0.125))-(Sincr, mvarform.VRuler.ScaleWidth)
                    mvarForm.VRuler.Line (mvarForm.VRuler.ScaleWidth - (mvarForm.VRuler.ScaleWidth * 0.125), Sincr)-(mvarForm.VRuler.ScaleWidth, Sincr)
                        End If
                    End If
                End If
            End If
            Sincr = Sincr + (RScale / RulerSize)
        Next
    Loop
End Sub
Public Sub ResizeControl(pControl As CommandButton)

'not used

Exit Sub

On Error Resume Next
    If mvarIsMoving Then
             With pControl
                .Width = mvarX
                '.Height = mvarY  ' don't care for height!
            End With
    End If
    
End Sub

Public Sub MoveControl(pControl As CommandButton)

'not used

Exit Sub

On Error GoTo errhandler:
    
        
    If mvarIsMoving Then
        
            With pControl
                .Move pControl.Left - mvarMoveX + mvarX, pControl.Top - mvarMoveY + mvarY
            End With
        
    End If
    
    mvarForm.Refresh
    
    'StatusBar1.Panels(2).Text = "X = " & TwipToInch((pControl.Left + mvarx) - piRulerOffsetX) & " Y = " & TwipToInch((pControl.Top + mvary) - piRulerOffsetY)
    'Debug.Print TwipToInch(pControl.Left), TwipToInch(CLng(mvarx)), TwipToInch((pControl.Left + mvarx) - piRulerOffsetX)
    'pControl.ToolTipText = " Left = " & TwipToInch(pControl.Left - piRulerOffsetX) & "  Top = " & TwipToInch(pControl.Top - piRulerOffsetY)
    'HandlePositionIndicator (pControl.Left + mvarx), (pControl.Top + mvary)
    
errhandler:
End Sub
Public Sub DrawGrid()
    Dim OldDrawMode As Integer, OldDrawSize As Integer, CX As Single, CY As Single
    Dim GridSpan As Integer, Row As Integer, Col As Integer
    
    mvarForm.Cls
    
    OldDrawMode = mvarForm.DrawMode
    
    mvarForm.DrawMode = vbInvert
    OldDrawSize = mvarForm.DrawWidth
    mvarForm.DrawWidth = 1
    
    GridSpan = CInt(GridSize * 1440)
    CY = GridSpan
    CX = GridSpan
    
    For Row = 1 To Int(mvarForm.ScaleHeight / GridSpan)
        DoEvents
        For Col = 1 To Int(mvarForm.ScaleWidth / GridSpan)
            mvarForm.PSet (CX, CY)
            CX = CX + GridSpan
        Next
        CX = GridSpan
        CY = CY + GridSpan
    Next
    
    mvarForm.DrawMode = OldDrawMode
    mvarForm.DrawWidth = OldDrawSize
    
End Sub
Public Property Let IsMoving(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsMoving = 5
    mvarIsMoving = vData
End Property


Public Property Get IsMoving() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsMoving
    IsMoving = mvarIsMoving
End Property



Public Property Let ResizeY(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ResizeY = 5
    mvarResizeY = vData
End Property


Public Property Get ResizeY() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ResizeY
    ResizeY = mvarResizeY
End Property



Public Property Let ResizeX(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ResizeX = 5
    mvarResizeX = vData
End Property


Public Property Get ResizeX() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ResizeX
    ResizeX = mvarResizeX
End Property



Public Property Let y(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Y = 5
    mvarY = vData
    'HandlePositionIndicator
End Property


Public Property Get y() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Y
    y = mvarY
End Property



Public Property Let x(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.X = 5
    mvarX = vData
    'HandlePositionIndicator
End Property


Public Property Get x() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.X
    x = mvarX
End Property



Public Property Let MoveY(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MoveY = 5
    mvarMoveY = vData
End Property


Public Property Get MoveY() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MoveY
    MoveY = mvarMoveY
End Property



Public Property Let MoveX(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MoveX = 5
    mvarMoveX = vData
End Property


Public Property Get MoveX() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MoveX
    MoveX = mvarMoveX
End Property



Public Sub HandlePositionIndicator(Optional pX As Single, Optional pY As Single)

On Error Resume Next

Dim iX As Single
Dim iY As Single

    If pX <> 0 Then
        iX = pX
    Else
        iX = mvarX
    End If
    
    If pY <> 0 Then
        iY = pY
    Else
        iY = mvarY
    End If
    
    
    
    mvarForm.HRuler.DrawMode = 6
    
    mvarForm.HRuler.Line (iX - mvarForm.HRuler.Left, 0)-(iX - mvarForm.HRuler.Left, mvarForm.HRuler.ScaleHeight), vbRed
    
    If mvarSaveX1 > 0 Then
        mvarForm.HRuler.Line (mvarSaveX1 - mvarForm.HRuler.Left, 0)-(mvarSaveX1 - mvarForm.HRuler.Left, mvarForm.HRuler.ScaleHeight)
    End If
    
    mvarForm.HRuler.DrawMode = 13
    mvarSaveX1 = iX
    
    mvarForm.VRuler.DrawMode = 6
    mvarForm.VRuler.Line (0, iY - mvarForm.VRuler.Top)-(mvarForm.VRuler.ScaleWidth, iY - mvarForm.VRuler.Top)
    If mvarSaveY1 > 0 Then
        mvarForm.VRuler.Line (0, mvarSaveY1 - mvarForm.VRuler.Top)-(mvarForm.VRuler.ScaleWidth, mvarSaveY1 - mvarForm.VRuler.Top)
    End If
    mvarForm.VRuler.DrawMode = 13
    mvarSaveY1 = iY
    
    
    mvarForm.sbStatusBar.Panels(3).Text = "Left = " & TwipToInch(iX - piRulerOffsetX) & " Top = " & TwipToInch(iY - piRulerOffsetY)
    
    'If TwipToInch(iX - piRulerOffsetX) < 0 Then
    '    mvarForm.mvarform.sbStatusBar.Panels(3).Text = "Left = Out of bounds" & " Top = " & TwipToInch(iY - piRulerOffsetY)
    '    Exit Sub
    'End If
    
    'If TwipToInch(iY - piRulerOffsetY) < 0 Then
    '    mvarForm.mvarform.sbStatusBar.Panels(3).Text = "Left = " & TwipToInch(iX - piRulerOffsetX) & " Top = Out of bounds"
    '    Exit Sub
    'End If
End Sub

Public Sub PicHandle_MouseDown(Index As Integer)
Dim i As Integer
Dim rc As RECT

    'Handles should only be visible when a control is selected
    Debug.Assert (Not Me.CurrCtl Is Nothing)

    'NOTE: ME.DragPointnot used for sizing
    'Save control position in screen coordinates

    Me.DragRect.SetRectToCtrl Me.CurrCtl
    Me.DragRect.TwipsToScreen Me.CurrCtl

    'Track index handle
    Me.DragHandle = Index

    'Hide sizing handles
    Me.ShowHandles False

    'We need to force handles to hide themselves before drawing drag rectangle
    mvarForm.Refresh

    'Indicate sizing is under way
    Me.DragState = StateSizing

    'Show sizing rectangle
    Me.DrawDragRect

    'In order to detect mouse movement over any part of the form,
    'we set the mouse capture to the form and will process mouse
    'movement from the applicable form events
    SetCapture mvarForm.hwnd

    'Limit cursor movement within form
    GetWindowRect mvarForm.hwnd, rc
    ClipCursor rc

End Sub

Private Sub Class_Initialize()
    If Not (frmMain Is Nothing) Then
        Set Me.MDIForm = frmMain
    End If
    
    Set mvarDragRect = New clsDragControls
    Set mvarDragPoint = New clsPoint
End Sub


Private Sub Class_Terminate()
    If Not (mvarForm Is Nothing) Then
         Set mvarForm = Nothing
    End If
    
    If Not (mvarMDIForm Is Nothing) Then
        Set mvarMDIForm = Nothing
    End If
    
    If Not (mvarDragRect Is Nothing) Then
        Set mvarDragRect = Nothing
    End If
    
    If Not (mvarDragPoint Is Nothing) Then
        Set mvarDragPoint = Nothing
    End If
    
End Sub
