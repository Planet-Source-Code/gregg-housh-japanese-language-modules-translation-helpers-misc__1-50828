Attribute VB_Name = "modJapaneseDates"
Option Explicit

' modJapaneseDates 1.0
' This module was written by Gregg Housh
'
' All the functions in here take a date variable
' as their only input, and return a string.

Public Function jWeekDay(ByVal TheDate As Date) As String
    On Error GoTo Err_Handler

    Dim strTemp As String
    
    Select Case Weekday(TheDate)
        Case 1 ' sunday
            strTemp = "Nichiyoobi"
        Case 2 ' monday
            strTemp = "Getsuyoobi"
        Case 3 ' tuesday
            strTemp = "Kayoobi"
        Case 4 ' wednesday
            strTemp = "Suiyoobi"
        Case 5 ' thursday
            strTemp = "Mokuyoobi"
        Case 6 ' friday
            strTemp = "Kinyoobi"
        Case 7 ' saturday
            strTemp = "Doyoobi"
    End Select
    
Exit_Handler:
    jWeekDay = strTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseDates", "jWeekDay")
    Err.Raise Err.Number
    GoTo Exit_Handler

End Function

Public Function jDay(ByVal TheDate As Date) As String
    On Error GoTo Err_Handler

    Dim strTemp As String
    
    Select Case Day(TheDate)
        Case 1
            strTemp = "tsuitachi"
        Case 2
            strTemp = "futsuka"
        Case 3
            strTemp = "mikka"
        Case 4
            strTemp = "yokka"
        Case 5
            strTemp = "itsuka"
        Case 6
            strTemp = "muika"
        Case 7
            strTemp = "nanoka"
        Case 8
            strTemp = "yooka"
        Case 9
            strTemp = "kokonoka"
        Case 10
            strTemp = "tooka"
        Case 11
            strTemp = "juuichi nichi"
        Case 12
            strTemp = "juuni nichi"
        Case 13
            strTemp = "juusan nichi"
        Case 14
            strTemp = "juuyokka"
        Case 15
            strTemp = "juugo nichi"
        Case 16
            strTemp = "juuroku nichi"
        Case 17
            strTemp = "juushichi nichi"
        Case 18
            strTemp = "juuhachi nichi"
        Case 19
            strTemp = "juuku nichi"
        Case 20
            strTemp = "hatsuka"
        Case 21
            strTemp = "nijuu ichi nichi"
        Case 22
            strTemp = "nijuu ni nichi"
        Case 23
            strTemp = "nijuu san nichi"
        Case 24
            strTemp = "nijuu yokka"
        Case 25
            strTemp = "nijuu go nichi"
        Case 26
            strTemp = "nijuu roku nichi"
        Case 27
            strTemp = "nijuu shichi nichi"
        Case 28
            strTemp = "nijuu hachi nichi"
        Case 29
            strTemp = "nijuu ku nichi"
        Case 30
            strTemp = "sanjuu nichi"
        Case 31
            strTemp = "sanjuu ichi nichi"
    End Select
Exit_Handler:
    jDay = strTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseDates", "jDay")
    Err.Raise Err.Number
    GoTo Exit_Handler

End Function

Public Function jMonth(ByVal TheDate As Date) As String
    On Error GoTo Err_Handler

    Dim strTemp As String
    
    Select Case Month(TheDate)
        Case 1
            strTemp = "ichi gatsu"
        Case 2
            strTemp = "ni gatsu"
        Case 3
            strTemp = "san gatsu"
        Case 4
            strTemp = "shi gatsu"
        Case 5
            strTemp = "go gatsu"
        Case 6
            strTemp = "roku gatsu"
        Case 7
            strTemp = "shichi gatsu"
        Case 8
            strTemp = "hachi gatsu"
        Case 9
            strTemp = "ku gatsu"
        Case 10
            strTemp = "juu gatsu"
        Case 11
            strTemp = "juu ichi gatsu"
        Case 12
            strTemp = "juu ni gatsu"
    End Select
Exit_Handler:
    jMonth = strTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseDates", "jMonth")
    Err.Raise Err.Number
    GoTo Exit_Handler

End Function

Public Function jHour(ByVal TheDate As Date) As String
    On Error GoTo Err_Handler

    Dim strTemp As String
    Dim intHour As Integer
    
    intHour = Hour(TheDate)
    
    If intHour > 12 Then
        strTemp = "gogo "
        intHour = intHour - 12
    Else
        strTemp = "gozen "
    End If
    
    Select Case intHour
        Case 1
            strTemp = strTemp & "ichiji"
        Case 2
            strTemp = strTemp & "niji"
        Case 3
            strTemp = strTemp & "sanji"
        Case 4
            strTemp = strTemp & "yoji"
        Case 5
            strTemp = strTemp & "goji"
        Case 6
            strTemp = strTemp & "rokuji"
        Case 7
            strTemp = strTemp & "shichiji"
        Case 8
            strTemp = strTemp & "hachiji"
        Case 9
            strTemp = strTemp & "kuji"
        Case 10
            strTemp = strTemp & "juuji"
        Case 11
            strTemp = strTemp & "juu ichiji"
        Case 12, 0
            strTemp = strTemp & "juu niji"
    End Select
    
Exit_Handler:
    jHour = strTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseDates", "jHour")
    Err.Raise Err.Number
    GoTo Exit_Handler

End Function

Public Function jMinute(ByVal TheDate As Date) As String
    On Error GoTo Err_Handler

    Dim strTemp As String
    Dim intMinute As Integer
    
    intMinute = Minute(TheDate)
    If intMinute < 11 Then
    
        Select Case intMinute
            Case 1
                strTemp = "ippun"
            Case 2
                strTemp = "nifun"
            Case 3
                strTemp = "sanpun"
            Case 4
                strTemp = "yonpun"
            Case 5
                strTemp = "gofun"
            Case 6
                strTemp = "roppun"
            Case 7
                strTemp = "nanafun"
            Case 8
                strTemp = "happun"
            Case 9
                strTemp = "kyuufun"
            Case 10
                strTemp = "juppun"
        End Select
    ElseIf intMinute >= 11 And intMinute <= 19 Then
        Select Case intMinute
            Case 11
                strTemp = "juuippun"
            Case 12
                strTemp = "juunifun"
            Case 13
                strTemp = "juusanpun"
            Case 14
                strTemp = "juuyonpun"
            Case 15
                strTemp = "juugofun"
            Case 16
                strTemp = "juuroppun"
            Case 17
                strTemp = "juunanafun"
            Case 18
                strTemp = "juuhappun"
            Case 19
                strTemp = "juukyuufun"
        End Select
    ElseIf intMinute >= 20 And intMinute <= 29 Then
        Select Case intMinute
            Case 20
                strTemp = "nijuppun"
            Case 21
                strTemp = "nijuuippun"
            Case 22
                strTemp = "nijuunifun"
            Case 23
                strTemp = "nijuusanpun"
            Case 24
                strTemp = "nijuuyonpun"
            Case 25
                strTemp = "nijuugofun"
            Case 26
                strTemp = "nijuuroppun"
            Case 27
                strTemp = "nijuunanafun"
            Case 28
                strTemp = "nijuuhappun"
            Case 29
                strTemp = "nijuukyuufun"
        End Select
    ElseIf intMinute >= 30 And intMinute <= 39 Then
        Select Case intMinute
            Case 30
                strTemp = "sanjuppun"
            Case 31
                strTemp = "sanjuuippun"
            Case 32
                strTemp = "sanjuunifun"
            Case 33
                strTemp = "sanjuusanpun"
            Case 34
                strTemp = "sanjuuyonpun"
            Case 35
                strTemp = "sanjuugofun"
            Case 36
                strTemp = "sanjuuroppun"
            Case 37
                strTemp = "sanjuunanafun"
            Case 38
                strTemp = "sanjuuhappun"
            Case 39
                strTemp = "sanjuukyuufun"
        End Select
    ElseIf intMinute >= 40 And intMinute <= 49 Then
        Select Case intMinute
            Case 40
                strTemp = "yonjuppun"
            Case 41
                strTemp = "yonjuuippun"
            Case 42
                strTemp = "yonjuunifun"
            Case 43
                strTemp = "yonjuuyonpun"
            Case 44
                strTemp = "yonjuuyonpun"
            Case 45
                strTemp = "yonjuugofun"
            Case 46
                strTemp = "yonjuuroppun"
            Case 47
                strTemp = "yonjuunanafun"
            Case 48
                strTemp = "yonjuuhappun"
            Case 49
                strTemp = "yonjuukyuufun"
        End Select
    ElseIf intMinute >= 50 And intMinute <= 59 Then
        Select Case intMinute
            Case 50
                strTemp = "gojuppun"
            Case 51
                strTemp = "gojuuippun"
            Case 52
                strTemp = "gojuunifun"
            Case 53
                strTemp = "gojuugopun"
            Case 54
                strTemp = "gojuuyonpun"
            Case 55
                strTemp = "gojuugofun"
            Case 56
                strTemp = "gojuuroppun"
            Case 57
                strTemp = "gojuunanafun"
            Case 58
                strTemp = "gojuuhappun"
            Case 59
                strTemp = "gojuukyuufun"
        End Select
    End If

Exit_Handler:
    jMinute = strTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseDates", "jMinute")
    Err.Raise Err.Number
    GoTo Exit_Handler

End Function

Public Function jHoliday(ByVal TheDate As Date) As String
    On Error GoTo Err_Handler

    Dim strTemp As String
    
    Select Case Month(TheDate)
        Case 1
            If Day(TheDate) = 1 Then
                strTemp = "Oshoogatsu"
            Else
                strTemp = "not a holiday"
            End If
        Case 3
            If Day(TheDate) = 3 Then
                strTemp = "Hinamatsuri"
            ElseIf Day(TheDate) > 20 Then
                strTemp = "Hanami"
            Else
                strTemp = "not a holiday"
            End If
        Case 4
            If Day(TheDate) < 10 Then
                strTemp = "Hanami"
            Else
                strTemp = "not a holiday"
            End If
        Case 5
            If Day(TheDate) = 5 Then
                strTemp = "Kodomo no Hi"
            Else
                strTemp = "not a holiday"
            End If
        Case 7
            If Day(TheDate) = 7 Then
                strTemp = "Tanabata"
            Else
                strTemp = "not a holiday"
            End If
        Case 8
            If Day(TheDate) > 20 Then
                strTemp = "Obon"
            Else
                strTemp = "not a holiday"
            End If
        Case 11
            If Day(TheDate) = 3 Then
                strTemp = "Bunka no Hi"
            Else
                strTemp = "not a holiday"
            End If
        Case 12
            If Day(TheDate) = 23 Then
                strTemp = "Tennoo Tanjoobi"
            ElseIf Day(TheDate) = 25 Then
                strTemp = "Kurisumasu"
            ElseIf Day(TheDate) = 31 Then
                strTemp = "Oomisoka"
            Else
                strTemp = "not a holiday"
            End If
        Case Else
            strTemp = "not a holiday"
    End Select
    
Exit_Handler:
    jHoliday = strTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseDates", "jHoliday")
    Err.Raise Err.Number
    GoTo Exit_Handler

End Function

