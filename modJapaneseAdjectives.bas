Attribute VB_Name = "modJapaneseAdjectives"
Option Explicit

' modJapaneseAdjectives 1.0
' This module was written by Gregg Housh
'
' WARNING:
' This module deals with true japanese adjectives.  It
' does not do anything with pseudo-adjectives (usually
' borrowed from chinese.)

' the A_ at the beginning of all of the enum entries
' is so they dont conflict with the ones in the enmVerbForms
' enumeration.
Public Enum enmAdjectiveForms
    A_BasicStem = 0
    A_Present_Informal = 1
    A_Present_Normal = 2
    A_Present_polite = 3
    A_Present_very_polite = 4
    A_Present_Negative_Informal = 5
    A_Present_Negative_Normal = 6
    A_Present_Negative_Polite = 7
    A_Present_Negative_Very_Polite = 8
    A_Past_Informal = 9
    A_Past_Normal = 10
    A_Past_Polite = 11
    A_Past_Very_Polite = 12
    A_Past_negative_Informal = 13
    A_Past_Negative_Normal = 14
    A_Past_Negative_Polite = 15
    A_Past_Negative_Very_Polite = 16
    A_Probable_Informal = 17
    A_Probable_Normal = 18
    A_Probable_Polite = 19
    A_Probable_Very_Polite = 20
    A_Probable_Negative_Informal = 21
    A_Probable_Negative_Normal = 22
    A_Probable_Negative_Polite = 23
    A_Probable_Negative_Very_Polite = 24
    A_Adverbial = 25
    A_Adverbial_negative = 26
    A_Suspending = 27
    A_Suspending_negative = 28
    A_Conditional_Present = 29
    A_Conditional_Present_Negative = 30
    A_Conditional_Past = 31
    A_Conditional_Past_negative = 32
End Enum

Public gAdjectiveExamples() As String

Public Function ConjugateAdjective(ByVal strAdjective As String, ByVal intType As enmAdjectiveForms) As String
    On Error GoTo Err_Handler
    
    Dim strTemp As String
    Dim strTemp2 As String
    strAdjective = LCase(strAdjective)
    If IsAdjective(strAdjective) Then
        Select Case intType
            Case enmAdjectiveForms.A_BasicStem
                strTemp = Left(strAdjective, Len(strAdjective) - 1)
            Case enmAdjectiveForms.A_Present_Informal
                ' same as the dictionary form
                strTemp = strAdjective
            Case enmAdjectiveForms.A_Present_Normal
                strTemp = strAdjective & " desu"
            Case enmAdjectiveForms.A_Present_polite
                strTemp = strAdjective & " no desu"
            Case enmAdjectiveForms.A_Present_very_polite
                strTemp = GetSpecialForm(strAdjective) & " gozaimasu"
            Case enmAdjectiveForms.A_Present_Negative_Informal
                strTemp = ConjugateAdjective(strAdjective, A_Adverbial) & " nai"
            Case enmAdjectiveForms.A_Present_Negative_Normal
                strTemp = ConjugateAdjective(strAdjective, A_Adverbial) & " nai desu"
            Case enmAdjectiveForms.A_Present_Negative_Polite
                strTemp = ConjugateAdjective(strAdjective, A_Adverbial) & " arimasen"
            Case enmAdjectiveForms.A_Present_Negative_Very_Polite
                strTemp = GetSpecialForm(strAdjective) & " gozaimasen"
            Case enmAdjectiveForms.A_Past_Informal
                strTemp = ConjugateAdjective(strAdjective, A_BasicStem) & "katta"
            Case enmAdjectiveForms.A_Past_Normal
                strTemp = ConjugateAdjective(strAdjective, A_BasicStem) & "katta deshita"
            Case enmAdjectiveForms.A_Past_Polite
                strTemp = ConjugateAdjective(strAdjective, A_BasicStem) & "katta no deshita"
            Case enmAdjectiveForms.A_Past_Very_Polite
                strTemp = GetSpecialForm(strAdjective) & " gozaimashita"
            Case enmAdjectiveForms.A_Past_negative_Informal
                strTemp = ConjugateAdjective(strAdjective, A_Adverbial) & " nakatta"
            Case enmAdjectiveForms.A_Past_Negative_Normal
                strTemp = ConjugateAdjective(strAdjective, A_Adverbial) & " nakatta deshita"
            Case enmAdjectiveForms.A_Past_Negative_Polite
                strTemp = ConjugateAdjective(strAdjective, A_Adverbial) & " nakatta no deshita"
            Case enmAdjectiveForms.A_Past_Negative_Very_Polite
                strTemp = GetSpecialForm(strAdjective) & " gozaimashita"
            Case enmAdjectiveForms.A_Probable_Informal
                strTemp = ConjugateAdjective(strAdjective, A_BasicStem) & "karoo"
            Case enmAdjectiveForms.A_Probable_Normal
                strTemp = ConjugateAdjective(strAdjective, A_BasicStem) & "karoo deshoo"
            Case enmAdjectiveForms.A_Probable_Polite
                strTemp = ConjugateAdjective(strAdjective, A_BasicStem) & "karoo no deshoo"
            Case enmAdjectiveForms.A_Probable_Very_Polite
                strTemp = GetSpecialForm(strAdjective) & " gozaimashoo"
            Case enmAdjectiveForms.A_Probable_Negative_Informal
                strTemp = ConjugateAdjective(strAdjective, A_Adverbial) & " nakaroo"
            Case enmAdjectiveForms.A_Probable_Negative_Normal
                strTemp = ConjugateAdjective(strAdjective, A_Adverbial) & " nakaroo deshoo"
            Case enmAdjectiveForms.A_Probable_Negative_Polite
                strTemp = ConjugateAdjective(strAdjective, A_Adverbial) & " nakaroo no deshoo"
            Case enmAdjectiveForms.A_Probable_Negative_Very_Polite
                strTemp = GetSpecialForm(strAdjective) & " gozaimashoo"
            Case enmAdjectiveForms.A_Adverbial
                strTemp = Left(strAdjective, Len(strAdjective) - 1) & "ku"
            Case enmAdjectiveForms.A_Adverbial_negative
                strTemp = Left(strAdjective, Len(strAdjective) - 1) & "ku naku"
            Case enmAdjectiveForms.A_Suspending
                strTemp = ConjugateAdjective(strAdjective, A_Adverbial) & "te"
            Case enmAdjectiveForms.A_Suspending_negative
                strTemp = ConjugateAdjective(strAdjective, A_Adverbial) & " nakute"
            Case enmAdjectiveForms.A_Conditional_Present
                strTemp = ConjugateAdjective(strAdjective, A_BasicStem) & "kereba"
            Case enmAdjectiveForms.A_Conditional_Present_Negative
                strTemp = ConjugateAdjective(strAdjective, A_Adverbial) & " nakereba"
            Case enmAdjectiveForms.A_Conditional_Past
                strTemp = ConjugateAdjective(strAdjective, A_BasicStem) & "kattara"
            Case enmAdjectiveForms.A_Conditional_Past_negative
                strTemp = ConjugateAdjective(strAdjective, A_Adverbial) & " nakattara"
        End Select
    Else
        strTemp = "Not an adjective"
    End If
    
Exit_Handler:
    ConjugateAdjective = strTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseAdjectives", "ConjugateAdjectives")
    Err.Raise Err.Number
    GoTo Exit_Handler

End Function

Public Function IsAdjective(ByVal strAdjective As String) As Boolean
    On Error GoTo Err_Handler
    
    Dim blnTemp As Boolean
    Dim strTemp As String
    strTemp = LCase(Right(strAdjective, 2))
    
    If Len(strAdjective) > 2 And (strTemp = "ai" Or strTemp = "ii" Or strTemp = "oi" Or strTemp = "ui") Or strAdjective = "ii" Then
        blnTemp = True
    Else
        blnTemp = False
    End If

Exit_Handler:
    IsAdjective = blnTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseAdjectives", "IsAdjective")
    Err.Raise Err.Number
    GoTo Exit_Handler

End Function

Public Function GetSpecialForm(ByVal strAdjective As String) As String
    On Error GoTo Err_Handler
    
    Dim strTemp As String
    Dim strTemp2 As String
    
    strTemp2 = LCase(Right(strAdjective, 2))
    Select Case strTemp2
        Case "oi"
            strTemp = Left(strAdjective, Len(strAdjective) - 2) & "oo"
        Case "ai"
            strTemp = Left(strAdjective, Len(strAdjective) - 2) & "oo"
        Case "ii"
            strTemp = Left(strAdjective, Len(strAdjective) - 2) & "uu"
        Case "ui"
            strTemp = Left(strAdjective, Len(strAdjective) - 2) & "uu"
        Case Else
            strTemp = "ERROR"
    End Select

Exit_Handler:
    GetSpecialForm = strTemp
    Exit Function

Err_Handler:
'    Call AddToCallStack("modJapaneseAdjectives", "GetSpecialForm")
    Err.Raise Err.Number
    GoTo Exit_Handler

End Function
