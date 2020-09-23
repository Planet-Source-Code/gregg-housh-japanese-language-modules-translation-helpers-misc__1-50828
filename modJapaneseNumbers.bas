Attribute VB_Name = "modJapaneseNumbers"
Option Explicit

' modJapaneseNumbers 1.0
' This module was written by Gregg Housh

' jNumber handles numbers from 0 to 999,999,999
' it converts them from numeric form to romanji
' ex: jNumber(1001) = sen ichi
' ex: jNumber(3000000) = sanbyakuman
Public Function jNumber(ByVal lngNumber As Long) As String
    On Error GoTo Err_Handler

    Dim strTemp As String
    Dim strTemp2 As String
    
    ' first we get the simple 0-10 out of the way
    If lngNumber >= 0 And lngNumber < 11 Then
        strTemp = GetZeroToTen(lngNumber)
    Else
        ' here we start to handle the harder ones.
        strTemp2 = CStr(lngNumber)
        If Len(strTemp2) = 2 Then
            ' 11 through 99
            ' get the first part
            strTemp = GetZeroToTen(CLng(Left(strTemp2, 1)))
            ' if its 1 then we dont want it...
            If strTemp = "ichi" Then strTemp = ""
            ' add "juu" and the last digit to the string
            strTemp = strTemp & "juu"
            strTemp2 = GetZeroToTen(CLng(Right(strTemp2, 1)))
            If strTemp2 <> "zero" Then
                strTemp = strTemp & " " & strTemp2
            End If
        ElseIf Len(strTemp2) >= 3 And Len(strTemp2) <= 9 Then
            ' 100 - 999999999
            ' get the first part
            ' this takes the first digit of the number, and
            ' the length of the number (string length, not bytes)
            ' ex: 101 , 1 and 3
            strTemp = GetBeginning(CLng(Left(strTemp2, 1)), Len(strTemp2))
            If Right(strTemp2, Len(strTemp2) - 1) <> String$(Len(strTemp2) - 1, "00") Then
                strTemp = strTemp & " " & jNumber(CLng(Right(strTemp2, Len(strTemp2) - 1)))
            End If
        Else
            strTemp = "COMING LATER"
        End If
    End If
    
Exit_Handler:
    jNumber = strTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseNumbers", "jNumber")
    Err.Raise Err.Number
    GoTo Exit_Handler
End Function

Public Function GetBeginning(ByVal lngStart As Long, ByVal lngLength As Long) As String
    On Error GoTo Err_Handler

    Dim strTemp As String
    
    Select Case lngLength
        Case 3
            strTemp = GetHundredToNineHundred(lngStart)
        Case 4
            strTemp = GetThousandToNineThousand(lngStart)
        Case 5
            strTemp = GetTenThousandToNinetyThousand(lngStart)
        Case 6
            strTemp = GetHundredThousandToNineHundredThousand(lngStart)
        Case 7
            strTemp = GetOneMillionToNineMillion(lngStart)
        Case 8
            strTemp = GetTenMillionToNinetyNineMillion(lngStart)
        Case 9
            strTemp = GetOneHundredMillionToNineHundredMillion(lngStart)
    End Select
    
Exit_Handler:
    GetBeginning = strTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseNumbers", "GetBeginning")
    Err.Raise Err.Number
    GoTo Exit_Handler

End Function

Public Function GetZeroToTen(ByVal lngNumber As Long) As String
    On Error GoTo Err_Handler

    Dim strTemp As String

    Select Case lngNumber
        Case 0
            strTemp = "zero" ' alternative: "rei"
        Case 1
            strTemp = "ichi"
        Case 2
            strTemp = "ni"
        Case 3
            strTemp = "san"
        Case 4
            strTemp = "shi" ' alternative: "yon"
        Case 5
            strTemp = "go"
        Case 6
            strTemp = "roku"
        Case 7
            strTemp = "shichi" ' alternative: "nana"
        Case 8
            strTemp = "hachi"
        Case 9
            strTemp = "kyuu" ' alternative "ku"
        Case 10
            strTemp = "juu"
        Case Else
            strTemp = ""
    End Select

Exit_Handler:
    GetZeroToTen = strTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseNumbers", "GetZeroToTen")
    Err.Raise Err.Number
    GoTo Exit_Handler
End Function

Public Function GetHundredToNineHundred(ByVal lngNumber As Long) As String
    On Error GoTo Err_Handler

    Dim strTemp As String

    Select Case lngNumber
        Case 1
            strTemp = "hyaku"
        Case 2
            strTemp = "nihyaku"
        Case 3
            strTemp = "sanbyaku"
        Case 4
            strTemp = "yonhyaku"
        Case 5
            strTemp = "gohyaku"
        Case 6
            strTemp = "roppyaku"
        Case 7
            strTemp = "nanahyaku"
        Case 8
            strTemp = "happyaku"
        Case 9
            strTemp = "kyuuhyaku"
        Case Else
            strTemp = ""
    End Select

Exit_Handler:
    GetHundredToNineHundred = strTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseNumbers", "GetHundredToNineHundred")
    Err.Raise Err.Number
    GoTo Exit_Handler
End Function

Public Function GetThousandToNineThousand(ByVal lngNumber As Long) As String
    On Error GoTo Err_Handler

    Dim strTemp As String

    Select Case lngNumber
        Case 1
            strTemp = "sen"
        Case 2
            strTemp = "nisen"
        Case 3
            strTemp = "sanzen"
        Case 4
            strTemp = "yonsen"
        Case 5
            strTemp = "gosen"
        Case 6
            strTemp = "rokusen"
        Case 7
            strTemp = "nanasen"
        Case 8
            strTemp = "hassen"
        Case 9
            strTemp = "kyuusen"
        Case Else
            strTemp = ""
    End Select

Exit_Handler:
    GetThousandToNineThousand = strTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseNumbers", "GetThousandToNineThousand")
    Err.Raise Err.Number
    GoTo Exit_Handler
End Function

Public Function GetTenThousandToNinetyThousand(ByVal lngNumber As Long) As String
    On Error GoTo Err_Handler

    Dim strTemp As String

    Select Case lngNumber
        Case 1
            strTemp = "ichiman"
        Case 2
            strTemp = "niman"
        Case 3
            strTemp = "sanman"
        Case 4
            strTemp = "yonman"
        Case 5
            strTemp = "goman"
        Case 6
            strTemp = "rokuman"
        Case 7
            strTemp = "nanaman"
        Case 8
            strTemp = "hachiman"
        Case 9
            strTemp = "kyuuman"
        Case Else
            strTemp = ""
    End Select

Exit_Handler:
    GetTenThousandToNinetyThousand = strTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseNumbers", "GetTenThousandToNinetyThousand")
    Err.Raise Err.Number
    GoTo Exit_Handler
End Function

Public Function GetHundredThousandToNineHundredThousand(ByVal lngNumber As Long) As String
    On Error GoTo Err_Handler

    Dim strTemp As String
    Dim lngTemp As Long
    
    lngTemp = lngNumber * 10
        
    strTemp = jNumber(lngTemp) & "man"
    
Exit_Handler:
    GetHundredThousandToNineHundredThousand = strTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseNumbers", "GetHundredThousandToNineHundredThousand")
    Err.Raise Err.Number
    GoTo Exit_Handler
End Function

Public Function GetOneMillionToNineMillion(ByVal lngNumber As Long) As String
    On Error GoTo Err_Handler

    Dim strTemp As String
    Dim lngTemp As Long
    
    lngTemp = lngNumber * 100
        
    strTemp = jNumber(lngTemp) & "man"
    
Exit_Handler:
    GetOneMillionToNineMillion = strTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseNumbers", "GetOneMillionToNineMillion")
    Err.Raise Err.Number
    GoTo Exit_Handler
End Function

Public Function GetTenMillionToNinetyNineMillion(ByVal lngNumber As Long) As String
    On Error GoTo Err_Handler

    Dim strTemp As String
    Dim lngTemp As Long
    
    lngTemp = lngNumber * 1000
        
    strTemp = jNumber(lngTemp) & "man"
    
Exit_Handler:
    GetTenMillionToNinetyNineMillion = strTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseNumbers", "GetTenMillionToNinetyNineMillion")
    Err.Raise Err.Number
    GoTo Exit_Handler
End Function

Public Function GetOneHundredMillionToNineHundredMillion(ByVal lngNumber As Long) As String
    On Error GoTo Err_Handler

    Dim strTemp As String
    Dim lngTemp As Long
    
    strTemp = jNumber(lngNumber) & "oku"
    
Exit_Handler:
    GetOneHundredMillionToNineHundredMillion = strTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseNumbers", "GetOneHundredMillionToNineHundredMillion")
    Err.Raise Err.Number
    GoTo Exit_Handler
End Function

