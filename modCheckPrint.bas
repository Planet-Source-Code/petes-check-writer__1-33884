Attribute VB_Name = "modCheckPrint"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const WM_USER = &H400&
Const ACM_OPEN = WM_USER + 100&

Public Const piRulerOffsetX As Integer = 570
Public Const piRulerOffsetY As Integer = 495
Public Const psSeller As String = "Pete Sral"

Public Enum enmMDIToolbar
    tb_Main = 1
    tb_Design = 2
    tb_Write = 3
End Enum

Public Enum enmAnimations
    anm_CopyToPrinter = 101
    anm_Printing = 102
    anm_SmallFindComputer = 103
    anm_SmallFindFile = 104
    anm_SmallHourglass = 105
    anm_SmallPrint = 106
    anm_SmallWriteBook = 107
End Enum

Public Enum ReturnFileName
    r_DriveLetter = 0 '0 -- Returns the drive letter
    r_DirPath = 1 '1 -- Returns the directory path
    r_FileName = 2 '2 -- Returns the filename (without the extension)
    r_FileExt = 3 '3 -- Returns the extension
    r_FileNameExt = 4 '4 -- Returns filename and extenstion
    r_DriverLetterPath = 5 '5 -- Returns Driver letter and dir path
End Enum

Public p_clsOptions As clsOptions
Public Function HandleError(plErr As Long, psErrDesc As String) As Boolean

On Error Resume Next

Dim vbResponse As VbMsgBoxResult
Dim sSubject As String
Dim sBody As String

    'Tell the user an error has occured and
    '     ask whether they want to email the autho
    '     r
    vbResponse = MsgBox("Error Number: " & Trim$(str$(plErr)) & vbCrLf & "Description: " & psErrDesc & vbCrLf & "Would you like To inform the author of this error via e-mail?", vbYesNo + vbQuestion, "Error")


    If vbResponse = vbYes Then
        'They want to email the author
        sSubject = "ºErrorº"
        
        'Format Correctly
        sSubject = Replace(sSubject, "&", "%26")
        sSubject = Replace(sSubject, " ", "%20")
        sSubject = Replace(sSubject, vbCrLf, "%0D%0A")
    
        'Set the body
        sBody = "An Error occured." & vbCrLf & "Application Title: " & App.Title & vbCrLf & "Application Version: " & App.Major & "." & App.Minor & App.Revision & vbCrLf & "Application Executable Name: " & App.EXEName & vbCrLf & "Error Number: " & Trim$(str$(plErr)) & vbCrLf & "Error Description: " & psErrDesc & vbCrLf & "Time/Datestamp: " & Format(Date, "Long Date") & " at " & Format(Time, "Medium Time") & vbCrLf & vbCrLf & "Please explain in detail what you were trying To Do when the error occured:"
        sBody = Replace(sBody, "&", "%26")
        sBody = Replace(sBody, " ", "%20")
        sBody = Replace(sBody, vbCrLf, "%0D%0A")
    
        'Actually Email You
        Shell "start mailto:pete@pjs-inc.com?Subject=" & sSubject & "&Body=" & sBody, vbHide
    
        HandleError = True
    Else
        HandleError = False
    End If
    
End Function

Public Sub LoadWebPage(psWebPage As String, psForm As Form)
On Error GoTo LoadWebPage_Err       'Error Code Inserted: 01/11/2001 10:08:17 PM        ----------ERR--
Dim i As Integer
i = 40000
    If Len(psWebPage) <> 0 Then
        ShellExecute psForm.hwnd, "open", psWebPage, "", "", SW_SHOW
    Else
        Exit Sub
    End If
Exit Sub
LoadWebPage_Err:                'Error Code Inserted: 01/11/2001 10:08:17 PM        ----------ERR--
If Err <> 0 Then
    If HandleError(Err.Number, Err.Description) Then
        MsgBox "Someone will be in contact with you within 48 hours." & vbCrLf & "If the e-mail reaches its destination!", vbInformation, "Error Mailed Notification"
    End If
    Err.Clear
End If
End Sub

Public Function GetFileName(ByVal TempPath As String, ReturnType As ReturnFileName)


    Dim DriveLetter As String
    Dim DirPath As String
    Dim FName As String
    Dim Extension As String
    Dim PathLength As Integer
    Dim ThisLength As Integer
    Dim Offset As Integer
    Dim FileNameFound As Boolean


    

    DriveLetter = ""
    DirPath = ""
    FName = ""
    Extension = ""


    If Mid(TempPath, 2, 1) = ":" Then ' Find the drive letter.
        DriveLetter = Left(TempPath, 2)
        TempPath = Mid(TempPath, 3)
    End If

    PathLength = Len(TempPath)


    For Offset = PathLength To 1 Step -1 ' Find the Next delimiter.


        Select Case Mid(TempPath, Offset, 1)
            Case ".": ' This indicates either an extension or a . or a ..
            ThisLength = Len(TempPath) - Offset


            If ThisLength >= 1 Then ' Extension
                Extension = Mid(TempPath, Offset, ThisLength + 1)
            End If

            TempPath = Left(TempPath, Offset - 1)
            Case "\": ' This indicates a path delimiter.
            ThisLength = Len(TempPath) - Offset


            If ThisLength >= 1 Then ' Filename
                FName = Mid(TempPath, Offset + 1, ThisLength)
                TempPath = Left(TempPath, Offset)
                FileNameFound = True
                Exit For
            End If

            Case Else
        End Select

Next Offset



If FileNameFound = False Then
    FName = TempPath
Else
    DirPath = TempPath
End If



If ReturnType = 0 Then
    GetFileName = DriveLetter
ElseIf ReturnType = 1 Then
    GetFileName = DirPath
ElseIf ReturnType = 2 Then
    GetFileName = FName
ElseIf ReturnType = 3 Then
    GetFileName = Mid$(Extension, 2)
ElseIf ReturnType = 4 Then
    GetFileName = FName & Extension
ElseIf ReturnType = 5 Then
    GetFileName = DriveLetter & DirPath
End If

End Function

Public Function FileExists(psFileName As String) As Boolean

If Len(psFileName) = 0 Or Right(psFileName, 1) = "\" Then
  FileExists = False: Exit Function
End If

FileExists = (Dir(psFileName) <> "")

End Function


Public Sub SetFormPosition(pForm As Form, Optional pbSave As Boolean = False)

    If Not pbSave Then
        If GetSetting(App.ProductName, pForm.Name, "WindowState", pForm.WindowState) <> vbMaximized Then
            pForm.Left = GetSetting(App.ProductName, pForm.Name, "Left", pForm.Left)
            pForm.Top = GetSetting(App.ProductName, pForm.Name, "Top", pForm.Top)
            pForm.Width = GetSetting(App.ProductName, pForm.Name, "Width", pForm.Width)
            pForm.Height = GetSetting(App.ProductName, pForm.Name, "Height", pForm.Height)
            pForm.WindowState = GetSetting(App.ProductName, pForm.Name, "WindowState", pForm.WindowState)
        End If
        'pForm.Refresh
    Else
        SaveSetting App.ProductName, pForm.Name, "Left", pForm.Left
        SaveSetting App.ProductName, pForm.Name, "Top", pForm.Top
        SaveSetting App.ProductName, pForm.Name, "Width", pForm.Width
        SaveSetting App.ProductName, pForm.Name, "Height", pForm.Height
        SaveSetting App.ProductName, pForm.Name, "WindowState", IIf(pForm.WindowState <> 1, pForm.WindowState, 0)
    End If
    
End Sub

Public Sub ShowOptions()

    frmOptions.Show vbModal, frmMain
    
End Sub
Public Sub ShowMessage(pAnimation As enmAnimations, Optional psMsg As String, Optional psCaption As String)

    Load frmShowAnim
    frmShowAnim.Parameters pAnimation, psMsg, psCaption
    frmShowAnim.Show , frmMain
    frmShowAnim.Refresh
    
End Sub

Public Sub LoadResAVI(pCtrlAnimation As Animation, pEnumAnimType As enmAnimations, Optional pbSetAutoPlay As Boolean = True)
    
 On Error Resume Next
    
    pCtrlAnimation.Visible = True
    pCtrlAnimation.AutoPlay = pbSetAutoPlay
    
    SendMessage pCtrlAnimation.hwnd, ACM_OPEN, ByVal App.hInstance, ByVal pEnumAnimType
    DoEvents
    
End Sub

Public Sub ClearAnim(pCtrlAnimation As Animation, Optional pbAnimVisible As Boolean = True)

On Error Resume Next
    
    'clear previous animation
    With pCtrlAnimation
        .AutoPlay = False
        .Close
        .AutoPlay = True
    End With
    
    pCtrlAnimation.Visible = pbAnimVisible
    
End Sub
Public Function SelectPrinter() As Boolean

'On Error Resume Next

On Error GoTo SelectPrinter_Error

    Dim cPrinter As New clsCmDlg
    
    If Len(p_clsOptions.DefaultPrinter) = 0 Then
        cPrinter.flags = cdlPDDisablePrintToFile + cdlPDNoPageNums + cdlPDNoSelection
        cPrinter.CancelError = True
        cPrinter.ShowPrinter
        SelectPrinter = True
        Set cPrinter = Nothing
    Else
        p_clsOptions.SetDefaultPrinter
        SelectPrinter = True
    End If
Exit Function
SelectPrinter_Error:
    
    If Not (cPrinter Is Nothing) Then
        Set cPrinter = Nothing
    End If
    
    If Err = 5 Then
        MsgBox "Print cancelled!", vbExclamation
        Err.Clear
    Else
        MsgBox "Error #: " & str(Err) & vbCrLf & "Error Desc: " & Error & vbCrLf & "Selecting Printer", vbExclamation
    End If
    SelectPrinter = False
    
    
End Function
Public Sub SetMDIStatusBarMode(pMode As enmMode)

On Error Resume Next

    With frmMain.sbMain.Panels(2)
    
        Select Case pMode
            Case md_None '= 0
                .Picture = frmMain.imlIcons.ListImages("None").Picture
                .Text = " None"
            Case md_New '= 1
                .Picture = frmMain.imlIcons.ListImages("New").Picture
                .Text = " New"
            Case md_Edit '= 2
                .Picture = frmMain.imlIcons.ListImages("Edit").Picture
                .Text = " Edit/Design"
            Case md_View '= 3
                .Picture = frmMain.imlIcons.ListImages("View").Picture
                .Text = " View"
        End Select
        
    End With
End Sub
Public Sub ShowToolbar(pToolbar As enmMDIToolbar)

On Error Resume Next

    With frmMain
        .tlbMain.Visible = (pToolbar = tb_Main)
        .tbDesign.Visible = (pToolbar = tb_Design)
        .tbWrite.Visible = (pToolbar = tb_Write)
    End With
    
End Sub

Public Sub PrintSampleData()

    Dim sDesign As String
    Dim cOpen As New clsIniFile
    Dim sSampleFile As String
    
        sSampleFile = App.path & "\Sample.ckc"
        If Not FileExists(sSampleFile) Then
            Exit Sub
        End If
        
        ShowMessage anm_CopyToPrinter, "Printing sample data on sample check!", "Printing Check Data"
            
            
            If Not SelectPrinter() Then
                Unload frmShowAnim
                Exit Sub
            End If
            
            With cOpen
            
                .FullPath = sSampleFile
            
                Printer.Font = "Courier New"
                Printer.FontSize = 10
                Printer.ScaleMode = vbTwips
                
                '[AccountInfo]
                .SectionName = "AccountInfo"
                Printer.CurrentX = .GetString("Left", 100) - piRulerOffsetX
                Printer.CurrentY = .GetString("Top", 100) - piRulerOffsetY
                Printer.Print "Acct #:123-45-67890"
                
                '[Date]
                 .SectionName = "Date"
                Printer.CurrentX = .GetString("Left", 100) - piRulerOffsetX
                Printer.CurrentY = .GetString("Top", 100) - piRulerOffsetY
                Printer.Print Format(Now, "mm/dd/yyyy")
                
                '[PayTo]
                .SectionName = "PayTo"
                Printer.CurrentX = .GetString("Left", 100) - piRulerOffsetX
                Printer.CurrentY = .GetString("Top", 100) - piRulerOffsetY
                Printer.Print "Pete Sral"
                
                '[AmountValue]
                Printer.FontSize = 12
                Printer.FontBold = True
                .SectionName = "AmountValue"
                Printer.CurrentX = .GetString("Left", 100) - piRulerOffsetX
                Printer.CurrentY = .GetString("Top", 100) - piRulerOffsetY
                Printer.Print "10000.00" & String(10 - Len("10000.00"), "-")
                Printer.FontSize = 10
                Printer.FontBold = False
                
                '[AmountText]
                .SectionName = "AmountText"
                Printer.CurrentX = .GetString("Left", 100) - piRulerOffsetX
                Printer.CurrentY = .GetString("Top", 100) - piRulerOffsetY
                Printer.Print modNumToText.NumToWord("10000.00") & String(51 - Len(Trim$(Mid("10000.00", 1, 51))), "-") & IIf(Len("10000.00") > 50, "...", "") 'Mid(mvarForm.lblNumText.Caption, 1, 51) & String(51 - Len(Trim$(Mid(mvarForm.lblNumText.Caption, 1, 51))), "-") & IIf(Len(mvarForm.lblNumText.Caption) > 50, "...", "")
                
                '[Memo]
                .SectionName = "Memo"
                Printer.CurrentX = .GetString("Left", 100) - piRulerOffsetX
                Printer.CurrentY = .GetString("Top", 100) - piRulerOffsetY
                Printer.Print "This is sample memo data"
                
                '[Name]
                .SectionName = "Name"
                Printer.CurrentX = .GetString("Left", 100) - piRulerOffsetX
                Printer.CurrentY = .GetString("Top", 100) - piRulerOffsetY
                Printer.Print "Joe Smith"
                
                'Me.OpenFile = .FullPath
                Printer.EndDoc
                Unload frmShowAnim
            End With
            
            
        

End Sub
Public Sub PrintHolder(Optional pbPrintCheckImage As Boolean = False)

On Error GoTo PrintHolder_Error

With Printer
    
        ShowMessage anm_CopyToPrinter, IIf(pbPrintCheckImage = True, "Printing Sample Check/Holder", "Printing Check Holder"), "Printing..."
        
        If Not SelectPrinter() Then
            Unload frmShowAnim
            Exit Sub
        End If
        
        
        .DrawStyle = vbSolid
        .ScaleMode = vbInches
        ''.ColorMode = 1
        .CurrentX = 0
        .CurrentY = 0
        
         If pbPrintCheckImage Then
            Printer.PaintPicture LoadResPicture("SampleCheck", vbResBitmap), .CurrentX, .CurrentY
        End If
        
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
        
        
        SetPrintHolderInfo pbPrintCheckImage
        
        .CurrentX = 0
        .CurrentY = 0
        
       
        
        Printer.EndDoc
        Unload frmShowAnim
    
    End With
Exit Sub
PrintHolder_Error:
    Unload frmShowAnim
    MsgBox "Error #: " & str(Err) & vbCrLf & "Error Desc: " & Error & vbCrLf & "Printing Check Holder", vbExclamation
    
End Sub
Private Sub SetPrintHolderInfo(Optional pbPrintCheckImage As Boolean = False)

On Error Resume Next

    With Printer
        .CurrentX = 0
        .CurrentY = 3.25
        
        .FontName = "Arial"
        .FontSize = 16
        .FontBold = True
        .FontUnderline = True
        If pbPrintCheckImage Then
            Printer.Print "Check Sample/Holder" & vbCrLf
        Else
            Printer.Print "Check Holder" & vbCrLf
        End If
        
        .FontBold = False
        .FontUnderline = False
        
        'GPF on HP 1000C????
        .FontName = "Arial"
         .FontSize = 12
        Printer.Print "Use with printer: " & Printer.DeviceName & " ONLY!!!!" & vbCrLf
        
         .FontName = "Courier New"
        .FontSize = 8
        If pbPrintCheckImage Then
            .FontBold = True
            .FontUnderline = True
            Printer.Print "Instructions - Check Sample"
            .FontUnderline = False
            .FontBold = False
            Printer.Print "Use the check sample above to write and print a sample check."
            Printer.Print "The sample check can also be used to help you with the check design."
            Printer.Print ""
        End If
        
        .FontBold = True
        .FontUnderline = True
        Printer.Print "Instructions - Check Holder"
        .FontUnderline = False
        .FontBold = False
        Printer.Print "1. Cut the dotted lines.  Do not start before the solid line.  Start ** ON ** the solid line."
        Printer.Print "    Do not cut past the solid line.  Stop ** ON ** the solid line."
        Printer.Print ""
        Printer.Print "2. Place the four corners of the check into the slits in the check holder.  Place the"
        Printer.Print "    check holder into your printer.  If your printer has as straight pass through or optional"
        Printer.Print "    paper tray (use that).  Limit the amount of bending that may occur during printing!" & vbCrLf
        
        Printer.Print "    The check should not bubble or move while being held by the check holder!!!"
        Printer.Print "    If it does... please check that the slits are cut from line to line."
        Printer.Print ""
                
        Printer.Font = "Arial"
        Printer.FontSize = 8
        
        .FontBold = True
        .FontUnderline = True
        Printer.Print "Standard Software Disclaimer" & vbCrLf
        .FontUnderline = False
        .FontBold = False
        Printer.Print "While " & psSeller & " makes every effort to deliver high quality products, we do not guarantee"
        Printer.Print "that our products are free from defects.  Our software is provided 'as is,' and you use"
        Printer.Print "the software at your own risk. We make no warranties as to performance, merchantability, "
        Printer.Print "fitness for a particular purpose, or any other warranties whether expressed or implied.  "
        Printer.Print "No oral or written communication from or information provided by  " & psSeller & "  shall create a "
        Printer.Print "warranty. Under no circumstances shall  " & psSeller & "  be liable for direct, indirect, special, "
        Printer.Print "incidental, or consequential damages resulting from the use, misuse, or inability to use "
        Printer.Print "this software, even if  " & psSeller & "  has been advised of the possibility of such damages."
        Printer.Print vbCrLf & vbCrLf

        Printer.Print Chr(169) & Format(Now, "yyyy") & "  " & psSeller & " , All rights reserved"
        Printer.Print "Check Writer: Version " & App.Major & "." & App.Minor & "." & App.Revision
        Printer.Print ""
        Printer.Print "Thank you for using the Check Writer software! - Have a nice day!"
        
        .CurrentX = 0
        .CurrentY = 0
        .FontName = "Courier New"
    
    End With
End Sub

Public Sub ShowPrintCheck(Optional pbShowData As Boolean, Optional pbShowDesign As Boolean)

On Error Resume Next

    Load frmSelectPrintType
    frmSelectPrintType.InitForm pbShowData, pbShowDesign
    frmSelectPrintType.Show vbModal, frmMain
    
    Select Case frmSelectPrintType.Response
        Case chk_None '= 0
                
        Case chk_Holder '= 1
            PrintHolder
        Case chk_Sample '= 2
            PrintHolder True
        Case chk_Data '= 3
            frmCheckWrite.mnuCheckPrintCheck_Click
        Case chk_Design '= 4
            frmCheckRuler.mnuDesignPrintDesignSample_Click
    End Select
    
End Sub
