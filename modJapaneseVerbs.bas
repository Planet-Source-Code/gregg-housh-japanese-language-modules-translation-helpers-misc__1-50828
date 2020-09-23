Attribute VB_Name = "modJapaneseVerbs"
Option Explicit

' modJapaneseVerbs 1.0
' This module was written by Gregg Housh
'
' WARNING:
' I used a few books, some study tools, and the web
' to compile all the information I needed to write
' the ConjugateVerb function.  There are however
' some exceptions to all the rules.  I have accounted
' for most of them in this code.  But one cannot be
' done with this function.  The verb kiru has two
' meanings.  One of them is "to cut" the other is
' "to wear".  The one meaning "to cut" is a godan
' (u-droppping/consonant) verb, the other "to wear"
' is an ichidan (ru-dropping/vowel) verb.  This
' code is written to work with "to cut".  You will
' have to handle the other one on your own.  Or maybe
' you come up with a simple way to handle both in code...
' Part 2 to the warning:
' I do not handle the "suru" verbs correctly at all.
' These are verbs that came from china.  A good example
' is "chui suru", which means "to pay attention".  You
' can pass this to the function, but you wont like the
' results.  Instead, just pass suru, and put them together
' like this:
' verb = "chui " & ConjugateVerb("suru",Past_Polite)
' verb will be equal to "chui shimashita" which is correct.

Public Enum enmVerbForms
    BasicStem = 0
    CombiningStem = 1
    CombiningStem_Negative = 2
    Present = 3
    Present_Negative_Polite = 4
    Present_Negative_Informal = 5
    Past_Polite = 6
    Past_Informal = 7
    Past_Negative_Polite = 8
    Past_Negative_Informal = 9
    Probable = 10
    Probable_Polite = 11
    Probable_Negative_Polite = 12
    Probable_Negative_Informal = 13
    Probable_Daroo_Present_Informal = 14
    Probable_Daroo_Present_Negative_Informal = 15
    Probable_Daroo_Past_Informal = 16
    Probable_Daroo_Past_Negative_Informal = 17
    Probable_Deshoo_Present_Polite = 18
    Probable_Deshoo_Present_Negative_Polite = 19
    Probable_Deshoo_Past_Polite = 20
    Probable_Deshoo_Past_Negative_Polite = 21
    Desire = 22
    Desire_Negative = 23
    Participle_Present = 24
    Participle_Present_Negative = 25
    Participle_Progressive_Present_Informal = 26
    Participle_Progressive_Past_Informal = 27
    Participle_Progressive_Probable_Informal = 28
    Participle_Progressive_Probable_Informal_Daroo = 29
    Participle_Progressive_Present_Polite = 30
    Participle_Progressive_Past_Polite = 31
    Participle_Progressive_Probable_Polite = 32
    Participle_Progressive_Probable_Polite_deshoo = 33
    Request = 34
    Request_Negative = 35
    Passive = 36
    Causative = 37
End Enum

Public gVerbExamples(37) As String

Public Sub SetupVerbExamples()
    On Error GoTo Err_Handler
    
    gVerbExamples(0) = "tabe" & vbCrLf & "No example sentance for stems."
    gVerbExamples(1) = "tabe" & vbCrLf & "No example sentance for stems."
    gVerbExamples(2) = "tabe" & vbCrLf & "No example sentance for stems."
    
    ' Present
    gVerbExamples(3) = "Tabemasu" & vbCrLf & "I eat."
    'Present_Negative_Polite
    gVerbExamples(4) = "Tabemasen" & vbCrLf & "I don't eat."
    'Present_Negative_Informal
    gVerbExamples(5) = "Tabenai" & vbCrLf & "I don't eat."
    'Past_Polite
    gVerbExamples(6) = "Tabemashita" & vbCrLf & "I ate."
    'Past_Informal
    gVerbExamples(7) = "Tabeta" & vbCrLf & "I ate."
    'Past_Negative_Polite
    gVerbExamples(8) = "Tabemasen deshita" & vbCrLf & "I didn't eat."
    'Past_Negative_Informal
    gVerbExamples(9) = "Tabenakatta" & vbCrLf & "I didn't eat."
    'Probable
    gVerbExamples(10) = "Tabeyoo" & vbCrLf & "I will probably eat."
    'Probable_Polite
    gVerbExamples(11) = "Tabemashoo" & vbCrLf & "Let's eat."
    'Probable_Negative_Polite
    gVerbExamples(12) = "Tabemasen deshoo" & vbCrLf & "Probably did not eat."
    'Probable_Negative_Informal
    gVerbExamples(13) = "Tabenai daroo" & vbCrLf & "Probably did not eat."
    'Probable_Daroo_Present_Informal
    gVerbExamples(14) = "Taberu daroo" & vbCrLf & "Would probably eat."
    'Probable_Daroo_Present_Negative_Informal
    gVerbExamples(15) = "Tabenai daroo" & vbCrLf & "Would probably not eat."
    'Probable_Daroo_Past_Informal
    gVerbExamples(16) = "Tabeta daroo" & vbCrLf & "Probably ate."
    'Probable_Daroo_Past_Negative_Informal
    gVerbExamples(17) = "Tabenakatta daroo" & vbCrLf & "Probably did not eat."
    'Probable_Deshoo_Present_Polite
    gVerbExamples(18) = "Taberu deshoo" & vbCrLf & "Would probably eat."
    'Probable_Deshoo_Present_Negative_Polite
    gVerbExamples(19) = "Tabenai deshoo" & vbCrLf & "Would probably not eat."
    'Probable_Deshoo_Past_Polite
    gVerbExamples(20) = "Tabeta deshoo" & vbCrLf & "Probably ate."
    'Probable_Deshoo_Past_Negative_Polite
    gVerbExamples(21) = "Tabenakatta deshoo" & vbCrLf & "Probably did not eat."
    'Desire
    gVerbExamples(22) = "Tabetai desu" & vbCrLf & "I want to eat."
    'Desire_Negative
    gVerbExamples(23) = "Tabetaku arimasen" & vbCrLf & "I don't want to eat."
    'Participle_Present
    gVerbExamples(24) = "Tabete" & vbCrLf & "I am eating."
    'Participle_Present_Negative
    gVerbExamples(25) = "Tabenakute" & vbCrLf & "I am not eating."
    'Participle_Progressive_Present_Informal
    gVerbExamples(26) = "Tabete iru" & vbCrLf & "This is what I am eating."
    'Participle_Progressive_Past_Informal
    gVerbExamples(27) = "Tabete ita" & vbCrLf & "This is what I was eating."
    'Participle_Progressive_Probable_Informal
    gVerbExamples(28) = "Tabete iroo" & vbCrLf & "They are probably eating already."
    'Participle_Progressive_Probable_Informal_Daroo
    gVerbExamples(29) = "Tabete iru daroo" & vbCrLf & "They are probably eating already."
    'Participle_Progressive_Present_Polite
    gVerbExamples(30) = "Tabete imasu" & vbCrLf & "This is what I am eating."
    'Participle_Progressive_Past_Polite
    gVerbExamples(31) = "Tabete imashita" & vbCrLf & "This is what I was eating."
    'Participle_Progressive_Probable_Polite
    gVerbExamples(32) = "Tabete imashoo" & vbCrLf & "They are probably eating already."
    'Participle_Progressive_Probable_Polite_deshoo
    gVerbExamples(33) = "Tabete iru deshoo" & vbCrLf & "They are probably eating already."
    'Request
    gVerbExamples(34) = "Tabete kudasai" & vbCrLf & "Please eat."
    'Request_Negative
    gVerbExamples(35) = "Tabenaide kudasai" & vbCrLf & "Please don't eat."
    'Passive
    gVerbExamples(36) = "Taberareru" & vbCrLf & "To be eaten."
    'Causative
    gVerbExamples(37) = "Tabesaseru" & vbCrLf & "To make someone eat."
    
Exit_Handler:
    Exit Sub

Err_Handler:
'    Call AddToCallStack("modJapaneseVerbs", "SetupVerbExamples")
    Err.Raise Err.Number
    GoTo Exit_Handler
End Sub

Public Function ConjugateVerb(ByVal strVerb As String, ByVal intType As enmVerbForms) As String
    On Error GoTo Err_Handler
    
    Dim strBasicStem As String
    Dim strCombiningStem As String
    Dim strTemp As String
    Dim strTemp2 As String
    
    If Len(strVerb) > 1 Then
        If Right(strVerb, 1) = "u" Then
            'strBasicStem = GetBasicStem(strVerb)
            
            Select Case intType
                Case enmVerbForms.BasicStem
                    strTemp = GetBasicStem(strVerb)
                Case enmVerbForms.CombiningStem
                    strTemp = GetCombiningStem(strVerb)
                Case enmVerbForms.CombiningStem_Negative
                    strTemp = GetCombiningStem(strVerb, True)
                Case enmVerbForms.Present
                    strTemp = GetCombiningStem(strVerb) & "masu"
                Case enmVerbForms.Present_Negative_Polite
                    strTemp = GetCombiningStem(strVerb) & "masen"
                Case enmVerbForms.Present_Negative_Informal
                    strTemp = GetCombiningStem(strVerb, True) & "nai"
                Case enmVerbForms.Past_Polite
                    strTemp = GetCombiningStem(strVerb) & "mashita"
                Case enmVerbForms.Past_Informal
                    If IsIchiDan(strVerb) Then
                        strTemp = GetBasicStem(strVerb) & "ta"
                    Else
                        If strVerb = "kuru" Then
                            strTemp = "kita"
                        ElseIf strVerb = "suru" Then
                            strTemp = "shita"
                        Else
                            strTemp2 = GetBasicStem(strVerb)
                            If Len(strTemp2) > 2 Then
                                strTemp2 = Left(strTemp2, Len(strTemp2) - 1)
                            End If
                            Select Case Right(strVerb, 2)
                                Case "ku"
                                    strTemp = strTemp2 & "ita"
                                Case "gu"
                                    strTemp = strTemp2 & "ida"
                                Case "su"
                                    If Len(strVerb) > 2 Then
                                        If Right(strVerb, 3) = "tsu" Then
                                            strTemp = strTemp2 & "tta"
                                        Else
                                            strTemp = strTemp2 & "shita"
                                        End If
                                    Else
                                        strTemp = strTemp2 & "shita"
                                    End If
                                Case "ru"
                                    strTemp = strTemp2 & "tta"
                                Case "mu", "nu", "bu"
                                    strTemp = strTemp2 & "nda"
                                Case Else
                                    If Len(strTemp2) = 1 Then
                                        strTemp = strTemp2 & "tta"
                                    End If
                            End Select
                        End If
                    End If
                Case enmVerbForms.Past_Negative_Polite
                    strTemp = GetCombiningStem(strVerb) & "masen deshita"
                Case enmVerbForms.Past_Negative_Informal
                    strTemp = GetCombiningStem(strVerb, True) & "nakatta"
                    
                Case enmVerbForms.Probable
                    If IsIchiDan(strVerb) Then
                        strTemp = GetBasicStem(strVerb) & "yoo"
                    Else
                        If strVerb = "suru" Then
                            strTemp = "shiyoo"
                        ElseIf strVerb = "kuru" Then
                            strTemp = "koyoo"
                        Else
                            strTemp = Left(strVerb, Len(strVerb) - 1) & "oo"
                        End If
                    End If
                Case enmVerbForms.Probable_Polite
                    strTemp = GetCombiningStem(strVerb) & "mashoo"
                Case enmVerbForms.Probable_Negative_Polite
                    strTemp = GetCombiningStem(strVerb) & "masen deshoo"
                Case enmVerbForms.Probable_Negative_Informal
                    strTemp = ConjugateVerb(strVerb, Present_Negative_Informal) & " daroo"
                Case enmVerbForms.Probable_Daroo_Present_Informal
                    strTemp = strVerb & " daroo"
                Case enmVerbForms.Probable_Daroo_Present_Negative_Informal
                    strTemp = ConjugateVerb(strVerb, Present_Negative_Informal) & " daroo"
                Case enmVerbForms.Probable_Daroo_Past_Informal
                    strTemp = ConjugateVerb(strVerb, Past_Informal) & " daroo"
                Case enmVerbForms.Probable_Daroo_Past_Negative_Informal
                    strTemp = ConjugateVerb(strVerb, Past_Negative_Informal) & " daroo"
                Case enmVerbForms.Probable_Deshoo_Present_Polite
                    strTemp = strVerb & " deshoo"
                Case enmVerbForms.Probable_Deshoo_Present_Negative_Polite
                    strTemp = ConjugateVerb(strVerb, Present_Negative_Informal) & " deshoo"
                Case enmVerbForms.Probable_Deshoo_Past_Polite
                    strTemp = ConjugateVerb(strVerb, Past_Informal) & " deshoo"
                Case enmVerbForms.Probable_Deshoo_Past_Negative_Polite
                    strTemp = ConjugateVerb(strVerb, Past_Negative_Informal) & " deshoo"
                Case enmVerbForms.Desire
                    strTemp = GetCombiningStem(strVerb) & "tai desu"
                Case enmVerbForms.Desire_Negative
                    strTemp = GetCombiningStem(strVerb) & "taku arimasen"
                Case enmVerbForms.Participle_Present
                    strTemp2 = ConjugateVerb(strVerb, Past_Informal)
                    strTemp = Left(strTemp2, Len(strTemp2) - 1) & "e"
                Case enmVerbForms.Participle_Present_Negative
                    strTemp = GetCombiningStem(strVerb, True) & "nakute"
                Case enmVerbForms.Participle_Progressive_Present_Informal
                    strTemp = ConjugateVerb(strVerb, Participle_Present) & " iru"
                Case enmVerbForms.Participle_Progressive_Past_Informal
                    strTemp = ConjugateVerb(strVerb, Participle_Present) & " ita"
                Case enmVerbForms.Participle_Progressive_Probable_Informal
                    strTemp = ConjugateVerb(strVerb, Participle_Present) & " iroo"
                Case enmVerbForms.Participle_Progressive_Probable_Informal_Daroo
                    strTemp = ConjugateVerb(strVerb, Participle_Present) & " iru daroo"
                Case enmVerbForms.Participle_Progressive_Present_Polite
                    strTemp = ConjugateVerb(strVerb, Participle_Present) & " imasu"
                Case enmVerbForms.Participle_Progressive_Past_Polite
                    strTemp = ConjugateVerb(strVerb, Participle_Present) & " imashita"
                Case enmVerbForms.Participle_Progressive_Probable_Polite
                    strTemp = ConjugateVerb(strVerb, Participle_Present) & " imashoo"
                Case enmVerbForms.Participle_Progressive_Probable_Polite_deshoo
                    strTemp = ConjugateVerb(strVerb, Participle_Present) & " iru deshoo"
                Case enmVerbForms.Request
                    strTemp = ConjugateVerb(strVerb, Participle_Present) & " kudasai"
                Case enmVerbForms.Request_Negative
                    strTemp = ConjugateVerb(strVerb, Present_Negative_Informal) & "de kudasai"
                Case enmVerbForms.Passive
                    strTemp = ConjugateVerb(strVerb, CombiningStem_Negative) & IIf(IsIchiDan(strVerb), "rareru", "reru")
                Case enmVerbForms.Causative
                    strTemp = ConjugateVerb(strVerb, CombiningStem_Negative) & IIf(IsIchiDan(strVerb), "saseru", "seru")
                Case Else
                
            End Select
        
        Else
            strTemp = ""
        End If
    Else
        strTemp = ""
    End If
    
Exit_Handler:
    ConjugateVerb = strTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseVerbs", "ConjugateVerb")
    Err.Raise Err.Number
    GoTo Exit_Handler

End Function

Public Function GetBasicStem(ByVal strVerb As String) As String
    On Error GoTo Err_Handler
    
    Dim strTemp As String
    Dim bIchiDan As Boolean
    Dim bGoDan As Boolean
    
    strTemp = LCase(strVerb)
    ' is it long enough?
    If Len(strTemp) > 1 Then
        ' does it end in u?
        If Right(strTemp, 1) = "u" Then
            ' check for Ichi-Dan or Go-Dan
            bIchiDan = IsIchiDan(strTemp)
            bGoDan = Not bIchiDan
            
            ' make basic stem depending on group
            If bGoDan Then
                ' we have to check for "tsu" since it
                ' is a little different
                If Len(strTemp) > 3 Then
                    If Right(LCase(strTemp), 3) = "tsu" Then
                        strTemp = Left(strTemp, Len(strTemp) - 2)
                    Else
                        strTemp = Left(strTemp, Len(strTemp) - 1)
                    End If
                Else
                    strTemp = Left(strTemp, Len(strTemp) - 1)
                End If
            ElseIf bIchiDan Then
                strTemp = Left(strTemp, Len(strTemp) - 2)
            Else
                Err.Raise 10000, "GetBasicStem", strVerb & " is not a verb"
            End If
            
        Else
            Err.Raise 10000, "GetBasicStem", strVerb & " is not a verb"
        End If
    Else
        Err.Raise 10000, "GetBasicStem", strVerb & " is not a verb"
    End If
    
Exit_Handler:
    GetBasicStem = strTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseVerbs", "GetBasicStem")
    Err.Raise Err.Number
    GoTo Exit_Handler
End Function

Public Function GetCombiningStem(ByVal strVerb As String, Optional ByVal bNegative As Boolean = False) As String
    On Error GoTo Err_Handler
    
    Dim strTemp As String
    Dim bIchiDan As Boolean
    Dim bGoDan As Boolean
    strVerb = LCase(strVerb)
    strTemp = strVerb
    ' is it long enough?
    If Len(strTemp) > 1 Then
        ' does it end in u?
        If Right(strTemp, 1) = "u" Then
            ' check for Ichi-Dan or Go-Dan
            bIchiDan = IsIchiDan(strTemp)
            bGoDan = Not bIchiDan
            
            ' make basic stem depending on group
            If bGoDan Then
                ' we have to check for "tsu" since it
                ' is a little different
                If Len(strTemp) > 3 Then
                    If Right(strTemp, 3) = "tsu" Then
                        strTemp = Left(strTemp, Len(strTemp) - 2)
                    Else
                        strTemp = Left(strTemp, Len(strTemp) - 1)
                    End If
                Else
                    strTemp = Left(strTemp, Len(strTemp) - 1)
                End If
                               
                ' we have to fix a problem here
                ' there are 2 irregular verbs, and 5
                ' slightly irregular verbs when forming
                ' the combining stem. here we fix it.
                If strVerb = "kuru" Then
                    If bNegative Then
                        strTemp = "ko"
                    Else
                        strTemp = "ki"
                    End If
                ElseIf strVerb = "suru" Then
                    strTemp = "shi"
                ElseIf strVerb = "gozaru" Or strVerb = "irassharu" Or strVerb = "kudasaru" Or strVerb = "nasaru" Or strVerb = "ossharu" Then
                    strTemp = Left(strTemp, Len(strTemp) - 1) & "i"
                Else
                    ' make combining stem
                    If Right(strTemp, 1) = "t" Then
                        strTemp = Left(strTemp, Len(strTemp) - 1) & "cha"
                    ElseIf Right(strTemp, 1) = "s" Then
                        strTemp = Left(strTemp, Len(strTemp) - 1) & "sha"
                    Else
                        If bNegative Then
                            If Right(strVerb, 2) = "ou" Or Right(strVerb, 2) = "iu" Or Right(strVerb, 2) = "au" Then
                                strTemp = strTemp & "wa"
                            Else
                                strTemp = strTemp & "a"
                            End If
                        Else
                            strTemp = strTemp & "i"
                        End If
                    End If
                End If
            ElseIf bIchiDan Then
                strTemp = Left(strTemp, Len(strTemp) - 2)
            Else
                Err.Raise 10000, "GetCombiningStem", strVerb & " is not a verb"
            End If
            
        Else
            Err.Raise 10000, "GetCombiningStem", strVerb & " is not a verb"
        End If
    Else
        Err.Raise 10000, "GetCombiningStem", strVerb & " is not a verb"
    End If
        
Exit_Handler:
    GetCombiningStem = strTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseVerbs", "GetCombiningStem")
    Err.Raise Err.Number
    GoTo Exit_Handler
End Function


Public Function IsIchiDan(ByVal strVerb As String) As Boolean
    On Error GoTo Err_Handler

    Dim bTemp As Boolean
    
    If Len(strVerb) > 1 Then
        ' does it end in u?
        If Right(strVerb, 1) = "u" Then
            ' check for Ichi-Dan or Go-Dan
            If Len(strVerb) >= 3 Then
                If Right(strVerb, 3) = "eru" Or Right(strVerb, 3) = "iru" Then
                    If strVerb = "shiru" Or strVerb = "hairu" Or strVerb = "kaeru" Or strVerb = "kiru" Then
                        bTemp = False
                    Else
                        bTemp = True
                    End If
                Else
                    bTemp = False
                End If
            Else
                bTemp = False
            End If
        End If
    End If
    
Exit_Handler:
    IsIchiDan = bTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseVerbs", "IsIchiDan")
    Err.Raise Err.Number
    GoTo Exit_Handler
End Function
