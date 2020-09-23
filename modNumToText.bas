Attribute VB_Name = "modNumToText"
Option Explicit

Const Factor As Long = 100

Function NumToWord(numval)
    Dim NTW, NText, dollars, cents, NWord, totalcents As String
    Dim decplace, TotalSets, cnt, LDollHold As Integer
    ReDim NumParts(9) As String 'Array For Amount (sets of three)
    ReDim Place(9) As String 'Array containing place holders
    Dim LDoll As Integer 'Length of the Dollars Text Amount
    Place(2) = " Thousand " '
    Place(3) = " Million " 'Place holder names For money
    Place(4) = " Billion " 'amounts
    Place(5) = " Trillion " '
    NTW = "" 'Temp value For the Function
    NText = round_curr(numval) 'Roundup the cents To eliminate     cents gr 2
    NText = Trim(Str(NText)) 'String representation of amount
    decplace = InStr(Trim(NText), ".") 'Position of Decimal 0 If none


    dollars = Trim(Left(NText, IIf(decplace = 0, Len(numval), decplace - 1)))
    LDoll = Len(dollars)
    cents = Trim(Right(NText, IIf(decplace = 0, 0, Abs(decplace - Len(NText)))))


        If Len(cents) = 1 Then
            cents = cents & "0"
        End If


        If (LDoll Mod 3) = 0 Then
            TotalSets = (LDoll \ 3)
        Else
            TotalSets = (LDoll \ 3) + 1
        End If
        cnt = 1
        LDollHold = LDoll


        Do While LDoll > 0
            NumParts(cnt) = IIf(LDoll > 3, Right(dollars, 3), Trim(dollars))


            dollars = IIf(LDoll > 3, Left(dollars, (IIf(LDoll < 3, 3, LDoll)) - 3), "")
                LDoll = Len(dollars)
                cnt = cnt + 1
            Loop


            For cnt = TotalSets To 1 Step -1 'step through NumParts array
                NWord = GetWord(NumParts(cnt)) 'convert 1 element of                NumParts
                NTW = NTW & NWord 'concatenate it To temp                variable
                If NWord <> "" Then NTW = NTW & Place(cnt)
            Next cnt 'loop through


            If LDollHold > 0 Then
                'NTW = NTW & " DOLLARS and " 'concatenate text
                NTW = Trim$(NTW) & " and" 'concatenate text
            Else
                NTW = NTW & " # VOID # " 'concatenate text
            End If
            
            totalcents = GetTens(cents) 'Convert cents part to word
            If totalcents = "" Then cents = "00" 'Concat NO if cents=0
            NTW = Trim$(NTW) & " " & Trim$(cents) & "/100"   'Concat Dollars and Cents
            NumToWord = NTW 'Assign word value to function
End Function

'-------------------------------------------------------------------------------------------------------------------------------
'=======================================
'     =========================
' The following function converts a numb
'     er from 1 to 9 to text
'=======================================
'     =========================


Function GetDigit(Digit)


    Select Case Val(Digit)
        Case 1: GetDigit = "One" '
        Case 2: GetDigit = "Two" '
        Case 3: GetDigit = "Three" '
        Case 4: GetDigit = "Four" ' Assign a numeric word value
        Case 5: GetDigit = "Five" ' based on a Single digit.
        Case 6: GetDigit = "Six" '
        Case 7: GetDigit = "Seven" '
        Case 8: GetDigit = "Eight" '
        Case 9: GetDigit = "Nine" '
        Case Else: GetDigit = "" '
    End Select
End Function 'End function GetDigit - return To calling program
'-------------------------------------------------------------------------------------------------------------------------------
'=======================================
'     =========================
' The following function converts a numb
'     er from 10 to 99 to text
'=======================================
'     =========================


Function GetTens(tenstext)
    Dim GT As String
    GT = "" 'null out the temporary Function value


    If Val(Left(tenstext, 1)) = 1 Then ' If value between 10-19


        Select Case Val(tenstext)
            Case 10: GT = "Ten" '
            Case 11: GT = "Eleven" '
            Case 12: GT = "Twelve" '
            Case 13: GT = "Thirteen" ' Retrieve numeric word
            Case 14: GT = "Fourteen" ' value If between ten and
            Case 15: GT = "Fifteen" ' nineteen inclusive.
            Case 16: GT = "Sixteen" '
            Case 17: GT = "Seventeen" '
            Case 18: GT = "Eighteen" '
            Case 19: GT = "Nineteen" '
            Case Else
        End Select
    
Else ' If value between 20-99


    Select Case Val(Left(tenstext, 1))
        
        Case 2: GT = "Twenty " '
        Case 3: GT = "Thirty " '
        Case 4: GT = "Forty " '
        Case 5: GT = "Fifty " ' Retrieve value if it is
        Case 6: GT = "Sixty " ' divisible by ten
        Case 7: GT = "Seventy " ' excluding the value ten.
        Case 8: GT = "Eighty " '
        Case 9: GT = "Ninety " '
        Case Else
    End Select
GT = GT & GetDigit(Right(tenstext, 1)) 'Retrieve ones place
End If

GetTens = GT ' Assign Function return value.
End Function
'-----------------------------------------------------------------------------------------------------------
'=======================================
'     ==========================
' The following function converts a numb
'     er from 0 to 999 to text
'=======================================
'     ==========================


Function GetWord(NumText)
    Dim GW As String, x As Integer
    GW = "" 'null out temporary Function value


    If Val(NumText) > 0 Then


        For x = 1 To Len(NumText) 'loop the length of NumText times


            Select Case Len(NumText)
                Case 3:


                If Val(NumText) > 99 Then
                    GW = GetDigit(Left(NumText, 1)) & " Hundred "
                End If
                NumText = Right(NumText, 2)
                Case 2:
                GW = GW & GetTens(NumText)
                NumText = ""
                Case 1:
                GW = GetDigit(NumText)
                Case Else
            End Select
    Next x
End If
GetWord = GW 'assign Function return value
End Function 'End function GetWord - Return To calling program
'---------------------------------------------------------------------------------------------------------------


Function round_curr(currValue)
    '
    'This rounds any currency field
    '
    If Val(currValue) <> 0 Then
    round_curr = Int(Val(currValue) * Factor + 0.5) / Factor
    End If
End Function

