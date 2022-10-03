Attribute VB_Name = "password_generator"
Option Explicit

'========================================
'=                                      =
'=     SUBROUTINES FOR TESTING ONLY     =
'=                                      =
'========================================

Private Sub subTestGenAsString()
    Debug.Print funcGenerateSinglesAsString()
End Sub

Private Sub subTestElementFromString()
    Dim strData As String
    Dim i As Integer
    strData = funcGenerateSinglesAsString(5)
    'Debug.Print strData
    For i = 1 To 5
        Debug.Print i & ": " & funcFakeArrayElementFromString(strData, i - 1)
    Next i
End Sub

Private Sub subTestRankingsString()
    Dim strData As String
    Dim i As Integer
    strData = funcGenerateSinglesAsString(50)
    For i = 1 To 50
        Debug.Print i & ": " & funcFakeArrayElementFromString(strData, i - 1)
    Next i
    Debug.Print funcGenerateRankingString(strData)
End Sub

Private Sub subTestIsCharacter()
    Debug.Print "Special: @ => " & funcIsSpecialCharacter("@")
    Debug.Print "Special: a => " & funcIsSpecialCharacter("a")
    Debug.Print
    Debug.Print "Digit: 2 => " & funcIsDigit("2")
    Debug.Print "Digit: C => " & funcIsDigit("C")
    Debug.Print
    Debug.Print "Lower Case: f => " & funcIsLowerCase("f")
    Debug.Print "Lower Case: F => " & funcIsLowerCase("F")
    Debug.Print "Lower Case: # => " & funcIsLowerCase("#")
    Debug.Print
    Debug.Print "Upper Case: D => " & funcIsUpperCase("D")
    Debug.Print "Upper Case: d => " & funcIsUpperCase("d")
    Debug.Print "Upper Case: = => " & funcIsUpperCase("=")
    Debug.Print
    Debug.Print "Char is: T => " & funcIsCharType("T")
    Debug.Print "Char is: n => " & funcIsCharType("n")
    Debug.Print "Char is: % => " & funcIsCharType("%")
    Debug.Print "Char is: 8 => " & funcIsCharType("8")
    
End Sub

Private Sub TestReplaceCharacterByType()
    Dim strTestCharacters As String
    strTestCharacters = "BCDFGHJKLMNPQRSTVWXYZbcdfghjkmnpqrstvwxyz23456789!@#$%^&*()+=?;:"
    Dim i As Integer
    For i = 1 To Len(strTestCharacters)
        'Debug.Print Mid(strTestCharacters, i, 1)
        Debug.Print Mid(strTestCharacters, i, 1) & " => " & funcReplaceCharByType(Mid(strTestCharacters, i, 1))
    Next i
End Sub

Private Sub TestConformPwdRules_NoDuplicates()
    Debug.Print funcConformPwdToRules_Duplicates("A2b1@2c3$")
    Debug.Print funcConformPwdToRules_Duplicates("TzC2@2z2F!z(2KhG")
    Debug.Print funcConformPwdToRules_Duplicates("TzC2@2K@F!y(4KhG")
    Debug.Print funcConformPwdToRules_Duplicates("6w4)FzL!y@7Hc&D@")
    Debug.Print funcConformPwdToRules_Duplicates("6w4@Fz7!y@6Hc&D@")
    Debug.Print funcConformPwdToRules_Duplicates("y4%P9*Hg;2z?8v9G")
    Debug.Print funcConformPwdToRules_Duplicates("Tv%9Ct3P2!DyJ9B#")
    Debug.Print funcConformPwdToRules_Duplicates("gD@v9Ry!2?9;z6Sn")
    Debug.Print funcConformPwdToRules_Duplicates("aaaBBB666###")
    Debug.Print
End Sub

Sub TestConformPwdRules_Minimums()
    Debug.Print funcConformPwdToRules_Minimums("aB2$sK*")
    Debug.Print funcConformPwdToRules_Minimums("fz@WxkwP75G")
    Debug.Print funcConformPwdToRules_Minimums("2w!F^DGK@W")
    Debug.Print funcConformPwdToRules_Minimums("2w!x6&")
    Debug.Print funcConformPwdToRules_Minimums("234")
    Debug.Print
End Sub

'===========================================================
'=                                                         =
'=     THE PRINCIPAL PROCEDURE FOR GENERATING PASWORDS     =
'=                                                         =
'===========================================================

Public Sub GeneratePasswords()
    'These string lists are the characters used to generate passwords.
    'Each of these string lists is exactly 50 characters long.
    'Take note that to prevent ambiguity, I do not use: vowels, y, 1, 0

    Dim strUpperCase As String
    strUpperCase = "BCDFGHJKLMNPQRSTVWXZBCDFGHJKLMNPQRSTVWXZBCDFGHJKLM"
    Dim strLowerCase As String
    strLowerCase = "bcdfghjkmnpqrstvwxzbcdfghjkmnpqrstvwxzbcdfghjkmnpq"
    Dim strDigits As String
    strDigits = "42345678982345678962345678952345678972345678923579"
    Dim strSpecialCharacters As String
    strSpecialCharacters = "!@#$%^&*()+=?;:!@#$%^&*()+=?;:!@#$%^&*()+=?;:!@#$%"
    Dim strOverageCharacters As String
    strOverageCharacters = "WBzy@!CDxw2@PFGvr3t#Hs$4KJZ9pqr%LMN+mn?6kjm^P*QR8(STG&hgf?qb5ZY&Vc7=WX&c)df"

    'These are the input cells on the Generator worksheet:
    'Number of Characters E5
    '
    'Minimum Lower Case E6
    'Minimum Upper Case E7
    'Minimum Digits E8
    'Minimum Special Characters E9
    '
    'Number of Passwords E11 (default is 10)
    With ThisWorkbook.Worksheets("Generator")
        'If the totals of the minimums is higher than the  number of characters specified, then we need to increase
        'the value for the number of characters in the password.
        If (.Range("E6").Value + .Range("E7").Value + .Range("E8").Value + .Range("E9").Value) > .Range("E5").Value Then
            .Range("E5").Value = (.Range("E6").Value + .Range("E7").Value + .Range("E8").Value + .Range("E9").Value)
        End If
        'If the user does not enter a number for how many passwords to generate, we default it to 10.
        If .Range("E11").Value < 1 Then
            .Range("E11").Value = 10
        End If

        'Declare variables
        Dim i As Integer
        Dim iProgressive As Integer
        Dim iPwd As Integer
        Dim intWritePwd As Integer
        intWritePwd = 13
        Dim strRankings As String
        Dim sngThisRnd As Single
        Dim strThisSetOfCharacters As String
        Dim strPwd As String

        'Clear out any previous passwords that were generated
        '.Unprotect Password:="qwerty"
        For i = 14 To 23
            .Cells(i, 2).Value = ""
            .Cells(i, 5).Value = ""
        Next i
        '.Protect Password:="qwerty"
        
        'Now we loop the algorithm for each password.
        For iPwd = 1 To .Range("E11").Value
            'Initialize the position marker for this password.
            iProgressive = 0
            'Reset and then reaquire a new rankings list for the password. The rankings list will
            'determine which characters we are extracting from the individual character string lists.
            strRankings = ""
            strRankings = funcGenerateRankingString(funcGenerateSinglesAsString(.Range("E5").Value))
            'Reset the password holder for this loop
            strPwd = ""
            'Reset the list of characters being used for this password
            strThisSetOfCharacters = ""
            'Initialize the random number for this loop. I do this so that we are not always useing
            'the exact same character pattern progress for each word. Doing this allows me to alter
            'the order of character patterns based on a random input from the application. You will
            'then see the orders being changed based on the value of the last digit in the random
            'number generated.
            sngThisRnd = Rnd()

            If Right(CStr(sngThisRnd), 1) <= 2 Then
                'This order: UDLS
                'Upper case
                For i = 1 To .Range("E7").Value
                    iProgressive = iProgressive + 1
                    strThisSetOfCharacters = strThisSetOfCharacters & Mid(strUpperCase, funcFakeArrayElementFromString(strRankings, iProgressive - 1), 1)
                Next i
                'Digits
                For i = 1 To .Range("E8").Value
                    iProgressive = iProgressive + 1
                    strThisSetOfCharacters = strThisSetOfCharacters & Mid(strDigits, funcFakeArrayElementFromString(strRankings, iProgressive - 1), 1)
                Next i
                'Lower case
                For i = 1 To .Range("E6").Value
                    iProgressive = iProgressive + 1
                    strThisSetOfCharacters = strThisSetOfCharacters & Mid(strLowerCase, funcFakeArrayElementFromString(strRankings, iProgressive - 1), 1)
                Next i
                'Special characters
                For i = 1 To .Range("E9").Value
                    iProgressive = iProgressive + 1
                    strThisSetOfCharacters = strThisSetOfCharacters & Mid(strSpecialCharacters, funcFakeArrayElementFromString(strRankings, iProgressive - 1), 1)
                Next i
                'Overage characters
                For i = 1 To .Range("E5").Value - iProgressive
                    strThisSetOfCharacters = strThisSetOfCharacters & Mid(strOverageCharacters, funcFakeArrayElementFromString(strRankings, (iProgressive - 1) + i), 1)
                Next i
            ElseIf Right(CStr(sngThisRnd), 1) <= 5 Then
                'This order: LDSU
                'Lower case
                For i = 1 To .Range("E6").Value
                    iProgressive = iProgressive + 1
                    strThisSetOfCharacters = strThisSetOfCharacters & Mid(strLowerCase, funcFakeArrayElementFromString(strRankings, iProgressive - 1), 1)
                Next i
                'Digits
                For i = 1 To .Range("E8").Value
                    iProgressive = iProgressive + 1
                    strThisSetOfCharacters = strThisSetOfCharacters & Mid(strDigits, funcFakeArrayElementFromString(strRankings, iProgressive - 1), 1)
                Next i
                'Special characters
                For i = 1 To .Range("E9").Value
                    iProgressive = iProgressive + 1
                    strThisSetOfCharacters = strThisSetOfCharacters & Mid(strSpecialCharacters, funcFakeArrayElementFromString(strRankings, iProgressive - 1), 1)
                Next i
                'Upper case
                For i = 1 To .Range("E7").Value
                    iProgressive = iProgressive + 1
                    strThisSetOfCharacters = strThisSetOfCharacters & Mid(strUpperCase, funcFakeArrayElementFromString(strRankings, iProgressive - 1), 1)
                Next i
                'Overage characters
                For i = 1 To .Range("E5").Value - iProgressive
                    strThisSetOfCharacters = strThisSetOfCharacters & Mid(strOverageCharacters, funcFakeArrayElementFromString(strRankings, (iProgressive - 1) + i), 1)
                Next i
            ElseIf Right(CStr(sngThisRnd), 1) <= 7 Then
                'This order: USDL
                'Upper case
                For i = 1 To .Range("E7").Value
                    iProgressive = iProgressive + 1
                    strThisSetOfCharacters = strThisSetOfCharacters & Mid(strUpperCase, funcFakeArrayElementFromString(strRankings, iProgressive - 1), 1)
                Next i
                'Special characters
                For i = 1 To .Range("E9").Value
                    iProgressive = iProgressive + 1
                    strThisSetOfCharacters = strThisSetOfCharacters & Mid(strSpecialCharacters, funcFakeArrayElementFromString(strRankings, iProgressive - 1), 1)
                Next i
                'Digits
                For i = 1 To .Range("E8").Value
                    iProgressive = iProgressive + 1
                    strThisSetOfCharacters = strThisSetOfCharacters & Mid(strDigits, funcFakeArrayElementFromString(strRankings, iProgressive - 1), 1)
                Next i
                'Lower case
                For i = 1 To .Range("E6").Value
                    iProgressive = iProgressive + 1
                    strThisSetOfCharacters = strThisSetOfCharacters & Mid(strLowerCase, funcFakeArrayElementFromString(strRankings, iProgressive - 1), 1)
                Next i
                'Overage characters
                For i = 1 To .Range("E5").Value - iProgressive
                    strThisSetOfCharacters = strThisSetOfCharacters & Mid(strOverageCharacters, funcFakeArrayElementFromString(strRankings, (iProgressive - 1) + i), 1)
                Next i
            Else
                'This order: USLD
                'Upper case
                For i = 1 To .Range("E7").Value
                    iProgressive = iProgressive + 1
                    strThisSetOfCharacters = strThisSetOfCharacters & Mid(strUpperCase, funcFakeArrayElementFromString(strRankings, iProgressive - 1), 1)
                Next i
                'Special characters
                For i = 1 To .Range("E9").Value
                    iProgressive = iProgressive + 1
                    strThisSetOfCharacters = strThisSetOfCharacters & Mid(strSpecialCharacters, funcFakeArrayElementFromString(strRankings, iProgressive - 1), 1)
                Next i
                'Lower case
                For i = 1 To .Range("E6").Value
                    iProgressive = iProgressive + 1
                    strThisSetOfCharacters = strThisSetOfCharacters & Mid(strLowerCase, funcFakeArrayElementFromString(strRankings, iProgressive - 1), 1)
                Next i
                'Digits
                For i = 1 To .Range("E8").Value
                    iProgressive = iProgressive + 1
                    strThisSetOfCharacters = strThisSetOfCharacters & Mid(strDigits, funcFakeArrayElementFromString(strRankings, iProgressive - 1), 1)
                Next i
                'Overage characters
                For i = 1 To .Range("E5").Value - iProgressive
                    strThisSetOfCharacters = strThisSetOfCharacters & Mid(strOverageCharacters, funcFakeArrayElementFromString(strRankings, (iProgressive - 1) + i), 1)
                Next i
            End If

            'Grab a new rankings list to further randomize the order of the characters
            strRankings = ""
            strRankings = funcGenerateRankingString(funcGenerateSinglesAsString(.Range("E5").Value))
            'Concatenate the password
            For i = 1 To .Range("E5").Value
                strPwd = strPwd & Mid(strThisSetOfCharacters, funcFakeArrayElementFromString(strRankings, i - 1), 1)
            Next i

            intWritePwd = intWritePwd + 1
            '.Cells(intWritePwd, 2).Value = "'" & strPwd
            
            strPwd = funcConformPwdToRules_AdjacentTypes(strPwd)
            'First pass at duplicates
            strPwd = funcConformPwdToRules_Duplicates(strPwd)
            'Second pass at duplicates
            strPwd = funcConformPwdToRules_Duplicates(strPwd)
            'Finally minimums
            .Cells(intWritePwd, 2).Value = "'" & funcConformPwdToRules_Minimums(strPwd)
            .Cells(intWritePwd, 5).Value = "Len:" & Len(.Cells(intWritePwd, 2).Value) & "; U:" & funcCountCharacterByType(.Cells(intWritePwd, 2).Value, "u") & "; L:" & funcCountCharacterByType(.Cells(intWritePwd, 2).Value, "l") & "; D:" & funcCountCharacterByType(.Cells(intWritePwd, 2).Value, "d") & "; S:" & funcCountCharacterByType(.Cells(intWritePwd, 2).Value, "s")
            'Debug.Print strPwd
        Next iPwd
    End With
End Sub

'====================================================
'=                                                  =
'=     SUBROUTINE FOR GENERATING RANDOM NUMBERS     =
'=                                                  =
'====================================================

Public Function funcGenerateSinglesAsString(Optional ByVal intMax As Integer = 50) As String
    'This function uses the rnd() function to generate numbers between zero and 1. Unfortunately,
    'the rnd() function will sometimes generate a number with more than fifteen digits. Excel has
    'a native 15-digit precision. Any number with more than 15 digits (in decimal places) will be
    'converted to scientific notation (eg, 7.272351E-02 instead of 0.0727235131759481). We cannot,
    'in any meaningful way, convert that number in scientific notation into an unlimited single or
    'double. It's at this point we all wish that Excel had floating point data types.
    '
    'To deal with this, I do a quick conversion of the single to a string and then examine the string
    'for the letter "E" which appears with scientific notation. If the answer is true, then I just
    'skip that number and go on to the next one. This is also why I'm using a DO-WHILE loop and not
    'an overly complicated FOR loop.
    '
    'Declare the variables
    Dim i As Integer
    Dim counter As Integer
    Dim sngRnd As Single
    'Initialize the looping counter
    counter = 0
    'Loop to build the string
     Do While counter <= intMax - 1
       'Reinitialize the randomization
        Randomize
        'Generate a random number
        sngRnd = Left(Rnd(), 6)
        'Set a conditional for scientific notation as described above
        If InStr(CStr(sngRnd), "E") = 0 Then
            'Makes sure the number is actually less than 1. In the grand scheme of things
            'it probably doesn't matter, but I like to make things bend to my will. In this
            'case, Microsoft CLAIMS that the rnd() function will return a number that is
            'between zero and 1. However, my real world experience has shown that this is
            'not always true, though no one important will admit it. Therefore, I put in
            'this conditional to make sure it is less than 1. That is also why I'm using
            'a DO-WHILE loop instead of a FOR loop.
            If sngRnd < 1 Then
                'Since we have a number with no scientific notation and is truly between
                'zero and 1, we up the counter so that we don't exceed the specified number.
                counter = counter + 1
                'Since we are creating a string of comma separated values, we need to add
                'a comma, but we only do that if this is NOT the first number added to
                'the string.
                If Len(funcGenerateSinglesAsString) > 0 Then
                    funcGenerateSinglesAsString = funcGenerateSinglesAsString & ","
                End If
                'Now concatenate the generated number with the string
                funcGenerateSinglesAsString = funcGenerateSinglesAsString & Left(sngRnd, 6)
                'Debug.Print counter & ": " & Left(sngRnd, 6)
            End If
        End If
    Loop
End Function

'=======================================================
'=                                                     =
'=     SUBROUTINE FOR RANKING THE ARRAY OF NUMBERS     =
'=                                                     =
'=======================================================

Public Function funcGenerateRankingString(ByVal strRandomNumbers As String) As String
    funcGenerateRankingString = ""
    'Make sure that we have input data to work with
    If InStr(strRandomNumbers, ",") > 0 Then
        'Declare array to hold the input random numbers. Trying to do this with VBA arrays
        'is like beating your head on the wall, but without any relief when you stop. After
        'years of being fed up with this major weakness in VBA, I wrote a class module to
        'use as a Perl-style array ... an array handling that makes sense.

        Dim objRandomNumbers As New clsArray
        objRandomNumbers.SetArrayFromString content:=strRandomNumbers

        'Declare other variables we need
        Dim i1 As Integer
        Dim i2 As Integer
        Dim intLower As Integer
        'Start the heavy lifting
        For i1 = 0 To objRandomNumbers.LastIndex
            'We are going to examine each element in the array (number) and see how
            'many other numbers are lower than the examined number in the same array.
            'So, if we examine a number and six other numbers are lower, that means
            'the number examined is in the 6 + 1 position, or the 7th position. That
            'means if there are six numbers lower, the examined number ranks 7th.
            '
            'Reset the counter for lower numbers
            intLower = 0
            'Loop through the array again, only looking for numbers that are lower than
            'the number that we are examining.
            For i2 = 0 To objRandomNumbers.LastIndex
                If objRandomNumbers.Element(i2) < objRandomNumbers.Element(i1) Then
                    intLower = intLower + 1
                End If
            Next i2
            'We are generating a string of rankings, not trying to return an array (that's
            'just another nightmare in VBA). So we need to add commas to the string of
            'rankings BUT we only do this if it is NOT the first ranking. Below is just
            'one of a few different ways to do it.
            If Len(funcGenerateRankingString) > 0 Then
                funcGenerateRankingString = funcGenerateRankingString & ","
            End If
            'Now we add the current evaluated numbers ranking.
            funcGenerateRankingString = funcGenerateRankingString & intLower + 1
        Next i1
        'Debug.Print "final ranking string: " & funcGenerateRankingString
        
        'Destroy the object to free up memory ... because Excel is a memory hog.
        Set objRandomNumbers = Nothing
    End If
End Function

'=====================================================================
'=                                                                   =
'=     SUBROUTINE TO PLUCK AN INDEXED ELEMENT FROM A STRING LIST     =
'=                                                                   =
'=====================================================================

Public Function funcFakeArrayElementFromString(ByVal strCSV As String, ByVal intElement As Integer) As Variant
    'We've generated a list of random numbers as a strong concatenated with commas. Now we need to be
    'able to grab those numbers from the list like we are grabbing them from an array, meaning we grab
    'them by a pseudo index position
    funcFakeArrayElementFromString = ""
    If InStr(strCSV, ",") > 0 Then
        'Debug.Print strCSV
        
        Dim objArrCSV As New clsArray
        objArrCSV.SetArrayFromString content:=strCSV

        funcFakeArrayElementFromString = objArrCSV.Element(intElement)
        Set objArrCSV = Nothing
    End If
End Function


'====================================
'=                                  =
'=     isA SUBROUTINE UTILITIES     =
'=                                  =
'= https://www.w3schools.com/charsets/ref_html_ascii.asp#:~:text=The%20ASCII%20Character%20Set&text=ASCII%20is%20a%207%2Dbit,are%20all%20based%20on%20ASCII.
'====================================
'
'Identiying and typing individual characters is core to the functionality
'of the password generator. This collection of procedures make that far
'less stressful

Public Function funcIsUpperCase(ByVal strChar As String) As Boolean
    funcIsUpperCase = False
    If strChar <> "" Then
        If Asc(strChar) >= 65 And Asc(strChar) <= 90 Then
            funcIsUpperCase = True
        End If
    End If
End Function

Public Function funcIsLowerCase(ByVal strChar As String) As Boolean
    funcIsLowerCase = False
    If strChar <> "" Then
        If Asc(strChar) >= 97 And Asc(strChar) <= 122 Then
            funcIsLowerCase = True
        End If
    End If
End Function

Public Function funcIsDigit(ByVal strChar As String) As Boolean
    funcIsDigit = False
    If strChar <> "" Then
        If Asc(strChar) >= 48 And Asc(strChar) <= 57 Then
            funcIsDigit = True
        End If
    End If
End Function

Public Function funcIsSpecialCharacter(ByVal strChar As String) As Boolean
    funcIsSpecialCharacter = False
    If strChar <> "" Then
        If Asc(strChar) >= 33 And Asc(strChar) <= 47 Then
            funcIsSpecialCharacter = True
        ElseIf Asc(strChar) >= 58 And Asc(strChar) <= 64 Then
            funcIsSpecialCharacter = True
        ElseIf Asc(strChar) >= 91 And Asc(strChar) <= 96 Then
            funcIsSpecialCharacter = True
        ElseIf Asc(strChar) >= 123 And Asc(strChar) <= 126 Then
            funcIsSpecialCharacter = True
        End If
    End If
End Function

Public Function funcIsCharType(ByVal strChar As String) As String
    funcIsCharType = ""
    If strChar <> "" Then
        If Asc(strChar) >= 65 And Asc(strChar) <= 90 Then
            'Upper case
            funcIsCharType = "u"
        ElseIf Asc(strChar) >= 97 And Asc(strChar) <= 122 Then
            'Lower case
            funcIsCharType = "l"
        ElseIf Asc(strChar) >= 48 And Asc(strChar) <= 57 Then
            'Digit
            funcIsCharType = "d"
        Else
            'Special character
            funcIsCharType = "s"
        End If
    End If
End Function


'====================================================
'=                                                  =
'=     ENFORCE RULES ON THE GENERATED PASSWORDS     =
'=                                                  =
'====================================================
'   Minimum number of each character type
'   No two contiguous types
'   No repeat characters
'   No dictionary words (there are no vowels in the upper or lower case letter lists)

Public Function funcConformPwdToRulesGetNewType(ByVal strTypeGrouping As String) As String
    'This is a problem solving procedure. When you are looking at characters next to each
    'other and deciding on how to insert new characters, you need to make sure that the
    'new character insertion does not violate the established rules. This method will
    'look at groupings of rules and then let the caller know what type of character
    'can be inserted without violationg the rules.
    'TYPES:
    '   u = upper case
    '   l = lower case
    '   d = digits
    '   s = special characters
    
    funcConformPwdToRulesGetNewType = ""
    Select Case strTypeGrouping
        Case "ll"
            funcConformPwdToRulesGetNewType = "u"
        Case "lu"
            funcConformPwdToRulesGetNewType = "d"
        Case "ld"
            funcConformPwdToRulesGetNewType = "u"
        Case "ls"
            funcConformPwdToRulesGetNewType = "d"
        Case "uu"
            funcConformPwdToRulesGetNewType = "l"
        Case "ul"
            funcConformPwdToRulesGetNewType = "s"
        Case "ud"
            funcConformPwdToRulesGetNewType = "s"
        Case "us"
            funcConformPwdToRulesGetNewType = "l"
        Case "dd"
            funcConformPwdToRulesGetNewType = "s"
        Case "du"
            funcConformPwdToRulesGetNewType = "l"
        Case "dl"
            funcConformPwdToRulesGetNewType = "u"
        Case "ds"
            funcConformPwdToRulesGetNewType = "l"
        Case "ss"
            funcConformPwdToRulesGetNewType = "d"
        Case "su"
            funcConformPwdToRulesGetNewType = "d"
        Case "sl"
            funcConformPwdToRulesGetNewType = "u"
        Case "sd"
            funcConformPwdToRulesGetNewType = "l"

        Case "uld"
            funcConformPwdToRulesGetNewType = "s"
        Case "uls"
            funcConformPwdToRulesGetNewType = "d"
        Case "udl"
            funcConformPwdToRulesGetNewType = "s"
        Case "uds"
            funcConformPwdToRulesGetNewType = "l"
        Case "usd"
            funcConformPwdToRulesGetNewType = "l"
        Case "usl"
            funcConformPwdToRulesGetNewType = "d"
        Case "uuu"
            funcConformPwdToRulesGetNewType = "l"
        Case "ull"
            funcConformPwdToRulesGetNewType = "d"
        Case "udd"
            funcConformPwdToRulesGetNewType = "s"
        Case "uss"
            funcConformPwdToRulesGetNewType = "l"
        Case "uud"
            funcConformPwdToRulesGetNewType = "l"
        Case "uul"
            funcConformPwdToRulesGetNewType = "s"
        Case "uus"
            funcConformPwdToRulesGetNewType = "l"

        Case "lud"
            funcConformPwdToRulesGetNewType = "s"
        Case "lus"
            funcConformPwdToRulesGetNewType = "d"
        Case "ldu"
            funcConformPwdToRulesGetNewType = "s"
        Case "lds"
            funcConformPwdToRulesGetNewType = "u"
        Case "lsd"
            funcConformPwdToRulesGetNewType = "u"
        Case "lsu"
            funcConformPwdToRulesGetNewType = "d"
        Case "lll"
            funcConformPwdToRulesGetNewType = "u"
        Case "ldd"
            funcConformPwdToRulesGetNewType = "u"
        Case "luu"
            funcConformPwdToRulesGetNewType = "d"
        Case "lss"
            funcConformPwdToRulesGetNewType = "u"
        Case "lld"
            funcConformPwdToRulesGetNewType = "u"
        Case "llu"
            funcConformPwdToRulesGetNewType = "d"
        Case "lls"
            funcConformPwdToRulesGetNewType = "u"

        Case "dul"
            funcConformPwdToRulesGetNewType = "s"
        Case "dus"
            funcConformPwdToRulesGetNewType = "l"
        Case "dlu"
            funcConformPwdToRulesGetNewType = "s"
        Case "dls"
            funcConformPwdToRulesGetNewType = "u"
        Case "dsu"
            funcConformPwdToRulesGetNewType = "l"
        Case "dsl"
            funcConformPwdToRulesGetNewType = "u"
        Case "duu"
            funcConformPwdToRulesGetNewType = "s"
        Case "dll"
            funcConformPwdToRulesGetNewType = "u"
        Case "dss"
            funcConformPwdToRulesGetNewType = "l"
        Case "ddd"
            funcConformPwdToRulesGetNewType = "u"
        Case "dds"
            funcConformPwdToRulesGetNewType = "l"
        Case "ddl"
            funcConformPwdToRulesGetNewType = "s"
        Case "ddu"
            funcConformPwdToRulesGetNewType = "s"

        Case "sul"
            funcConformPwdToRulesGetNewType = "d"
        Case "sud"
            funcConformPwdToRulesGetNewType = "l"
        Case "slu"
            funcConformPwdToRulesGetNewType = "d"
        Case "sld"
            funcConformPwdToRulesGetNewType = "u"
        Case "sdu"
            funcConformPwdToRulesGetNewType = "l"
        Case "sdl"
            funcConformPwdToRulesGetNewType = "u"
        Case "sdd"
            funcConformPwdToRulesGetNewType = "l"
        Case "sll"
            funcConformPwdToRulesGetNewType = "d"
        Case "suu"
            funcConformPwdToRulesGetNewType = "l"
        Case "ssd"
            funcConformPwdToRulesGetNewType = "u"
        Case "ssl"
            funcConformPwdToRulesGetNewType = "u"
        Case "ssu"
            funcConformPwdToRulesGetNewType = "d"
        Case "sss"
            funcConformPwdToRulesGetNewType = "l"
        Case Else
            funcConformPwdToRulesGetNewType = "u"
    End Select
End Function

Public Function funcConformPwdToRules_AdjacentTypes(ByVal strPwd As String) As String
    'A good password does not have two characters of the same type (upper case, lower case,
    'digits, special characters) adjacent to each other. This procedure solves that problem.
    If strPwd <> "" Then
        'Declare variables
        Dim strThisTypeGroup As String
        Dim strNewType As String
        
        Dim strOriginalType As String
        Dim intOriginalPosition As Integer
        
        Dim intCountUpperCase As Integer
        Dim intCountLowerCase As Integer
        Dim intCountDigit As Integer
        Dim intCountSpecialCharacters As Integer
        Dim i As Integer
        
        'Declare and initialize the charcter lists
        Dim strUpperCase As String
        strUpperCase = "BCDFGHJKLMNPQRSTVWXZBCDFGHJKLMNPQRSTVWXZBCDFGHJKLM"
        Dim strLowerCase As String
        strLowerCase = "bcdfghjkmnpqrstvwxzbcdfghjkmnpqrstvwxzbcdfghjkmnpq"
        Dim strDigits As String
        strDigits = "42345678982345678962345678952345678972345678923579"
        Dim strSpecialCharacters As String
        strSpecialCharacters = "!@#$%^&*()+=?;:!@#$%^&*()+=?;:!@#$%^&*()+=?;:!@#$%"
        
        'Use a randomly generated integer to further randomize the character lists
        Dim sngRnd As Single
        sngRnd = Rnd()
        If Right(CStr(sngRnd), 1) >= 8 Then
            strUpperCase = "JKLMNPQRSTVWXZBCDFGHJKLMNPQRSTVWXZBCDFGHJKBCDFGHJK"
            strLowerCase = "ghjkmnpqrstvwxzbcdfghjkmnpqrstvwxzbcdfghjkmnbcdfgh"
            strDigits = "34567898234567896234567895234567897234567892357942"
            strSpecialCharacters = "#$%^&*()+=?;:!@#$%^&*()+=?;:!@#$%^&*()+=?;:!@#$%!@"
        
        ElseIf Right(CStr(sngRnd), 1) > 5 Then
            strUpperCase = "TVWXZBCDFGHJKLMNPQRSTVWXZBCDFGHJKBCDFGHJKLMNPQRSTV"
            strLowerCase = "mnpqrstvwxyzbcdfghjkmnpqrstvwxyzbcdfghjkmnbcdfghjk"
            strDigits = "67898234567896234567895234567897234567892357942345"
            strSpecialCharacters = "&*()+=?;:!@#$%^&*()+=?;:!@#$%^&*()+=?;:!@#$%!@#$%^"
        ElseIf Right(CStr(sngRnd), 1) > 3 Then
            strUpperCase = "WXZBCDFGHJKLMNPQRSTVWXZBCDFGHJKBCDFGHJKLMNPQRSTVWX"
            strLowerCase = "stvwxzbcdfghjkmnpqrstvwxzbcdfghjkmnbcdfghjkmnpqrst"
            strDigits = "89823456789623456789523456789723456789235794234567"
            strSpecialCharacters = ")+=?;:!@#$%^&*()+=?;:!@#$%^&*()+=?;:!@#$%!@#$%^&*("
        End If

        'Loop through all the characters in the password
        For i = 1 To Len(strPwd)
            'Reset variable to keep track of the iteration point in the character list for the character we are looking at replacing
            intOriginalPosition = 0
            'Reset string variables
            strThisTypeGroup = ""
            strNewType = ""
            If i = 1 Then
                'If it is the first letter in the pwd, then we do not need to look at the letter before it as there is none.
                If funcIsCharType(Left(strPwd, 1)) = funcIsCharType(Mid(strPwd, 2, 1)) Then
                    strOriginalType = funcIsCharType(Left(strPwd, 1))
                    If strOriginalType = "u" Then
                        intOriginalPosition = InStr(strUpperCase, Left(strPwd, 1))
                    ElseIf strOriginalType = "l" Then
                        intOriginalPosition = InStr(strLowerCase, Left(strPwd, 1))
                    ElseIf strOriginalType = "d" Then
                        intOriginalPosition = InStr(strDigits, Left(strPwd, 1))
                    Else
                        intOriginalPosition = InStr(strSpecialCharacters, Left(strPwd, 1))
                    End If
                    strThisTypeGroup = funcIsCharType(Mid(strPwd, 1, 1))
                    strThisTypeGroup = strThisTypeGroup & funcIsCharType(Mid(strPwd, 2, 1))
                    strNewType = funcConformPwdToRulesGetNewType(strThisTypeGroup)
                End If
            ElseIf i = Len(strPwd) Then
                'If is is the last letter in the pwd, then we do not need to look at the letter after it as there is none.
                If funcIsCharType(Right(strPwd, 1)) = funcIsCharType(Mid(strPwd, Len(strPwd) - 1, 1)) Then
                    strOriginalType = funcIsCharType(Right(strPwd, 1))
                    If strOriginalType = "u" Then
                        intOriginalPosition = InStr(strUpperCase, Right(strPwd, 1))
                    ElseIf strOriginalType = "l" Then
                        intOriginalPosition = InStr(strLowerCase, Right(strPwd, 1))
                    ElseIf strOriginalType = "d" Then
                        intOriginalPosition = InStr(strDigits, Right(strPwd, 1))
                    Else
                        intOriginalPosition = InStr(strSpecialCharacters, Right(strPwd, 1))
                    End If
                    strThisTypeGroup = funcIsCharType(Mid(strPwd, i - 1, 1))
                    strThisTypeGroup = strThisTypeGroup & funcIsCharType(strThisTypeGroup & Mid(strPwd, i, 1))
                    strNewType = funcConformPwdToRulesGetNewType(strThisTypeGroup)
                End If
            Else
                'For each letter that is not the first or last, we need to look at the letter before, the letter, the letter after.
                If funcIsCharType(Mid(strPwd, i - 1, 1)) = funcIsCharType(Mid(strPwd, i, 1)) Or funcIsCharType(Mid(strPwd, i + 1, Len(strPwd))) = funcIsCharType(Mid(strPwd, i, 1)) Then
                    strOriginalType = funcIsCharType(Mid(strPwd, i, 1))
                    If strOriginalType = "u" Then
                        intOriginalPosition = InStr(strUpperCase, Mid(strPwd, i, 1))
                    ElseIf strOriginalType = "l" Then
                        intOriginalPosition = InStr(strLowerCase, Mid(strPwd, i, 1))
                    ElseIf strOriginalType = "d" Then
                        intOriginalPosition = InStr(strDigits, Mid(strPwd, i, 1))
                    Else
                        intOriginalPosition = InStr(strSpecialCharacters, Mid(strPwd, i, 1))
                    End If
                    strThisTypeGroup = funcIsCharType(Mid(strPwd, i - 1, 1))
                    strThisTypeGroup = strThisTypeGroup & funcIsCharType(strThisTypeGroup & Mid(strPwd, i, 1))
                    strThisTypeGroup = strThisTypeGroup & funcIsCharType(Mid(strPwd, i + 1, 1))
                    strNewType = funcConformPwdToRulesGetNewType(strThisTypeGroup)
                End If
            End If
            'Debug.Print "char and pos and group type and type: " & Mid(strPwd, i, 1) & " " & intOriginalPosition & " " & strThisTypeGroup & " " & strNewType
            
            'If we have an original position detected AND a new type detected, then we need to make a replacement
            If intOriginalPosition > 0 And strNewType <> "" Then
                If strNewType = "u" Then
                    If i = 1 Then
                        'Replacing the first letter
                        strPwd = Mid(strUpperCase, 1, 1) & Mid(strPwd, 2, Len(strPwd))
                    ElseIf i = Len(strPwd) Then
                        'Replacing the last letter
                        strPwd = Mid(strPwd, 1, Len(strPwd) - 1) & Mid(strUpperCase, i, 1)
                    Else
                        'Replaceing a letter not at the ends
                        strPwd = Left(strPwd, i - 1) & Mid(strUpperCase, i, 1) & Mid(strPwd, i + 1, Len(strPwd))
                    End If
                ElseIf strNewType = "l" Then
                    If i = 1 Then
                        'Replacing the first letter
                        strPwd = Mid(strLowerCase, 1, 1) & Mid(strPwd, 2, Len(strPwd))
                    ElseIf i = Len(strPwd) Then
                        'Replacing the last letter
                        strPwd = Mid(strPwd, 1, Len(strPwd) - 1) & Mid(strLowerCase, i, 1)
                    Else
                        'Replaceing a letter not at the ends
                        strPwd = Left(strPwd, i - 1) & Mid(strLowerCase, i, 1) & Mid(strPwd, i + 1, Len(strPwd))
                    End If
                ElseIf strNewType = "d" Then
                    If i = 1 Then
                        'Replacing the first letter
                        strPwd = Mid(strDigits, 1, 1) & Mid(strPwd, 2, Len(strPwd))
                    ElseIf i = Len(strPwd) Then
                        'Replacing the last letter
                        strPwd = Mid(strPwd, 1, Len(strPwd) - 1) & Mid(strDigits, i, 1)
                    Else
                        'Replaceing a letter not at the ends
                        strPwd = Left(strPwd, i - 1) & Mid(strDigits, i, 1) & Mid(strPwd, i + 1, Len(strPwd))
                    End If
                Else
                    If i = 1 Then
                        'Replacing the first letter
                        strPwd = Mid(strSpecialCharacters, 1, 1) & Mid(strPwd, 2, Len(strPwd))
                    ElseIf i = Len(strPwd) Then
                        'Replacing the last letter
                        strPwd = Mid(strPwd, 1, Len(strPwd) - 1) & Mid(strSpecialCharacters, i, 1)
                    Else
                        'Replaceing a letter not at the ends
                        strPwd = Left(strPwd, i - 1) & Mid(strSpecialCharacters, i, 1) & Mid(strPwd, i + 1, Len(strPwd))
                    End If
                End If
            End If
        Next i
    End If
    funcConformPwdToRules_AdjacentTypes = strPwd
End Function

Public Function funcConformPwdToRules_Duplicates(ByVal strPwd As String) As String
    'In a password, we do not want to duplicate characters. Having all unique characters
    'makes a password take longer to crack. This procedure will remove a duplicate
    'character and replace it with a new one of the same type. NOTE: This procedure is
    'only 99% effective. Sometimes a double will slip through.
    
    funcConformPwdToRules_Duplicates = strPwd
    'Declare and initialize variables
    Dim i As Integer
    Dim i2 As Integer
    Dim strNewPwd As String

    'Declare and initialize counters
    Dim incrementUpper As Integer
    incrementUpper = 0
    Dim incrementLower As Integer
    incrementLower = 0
    Dim incrementDigit As Integer
    incrementDigit = 0
    Dim incrementSpecial As Integer
    incrementSpecial = 0

    'Declare and initialize the character lists
    Dim strUpperCase As String
    strUpperCase = "KJHGFDCBZXWVTSRQPNMLKJHGFDCBZXWVTSRQPNMLKJHGFDCBZX"
    Dim strLowerCase As String
    strLowerCase = "jhgfdcbzxwvtsrqpnmkjhgfdcbzxwvtsrqpnmkjhgfdcbzyxwv"
    Dim strDigits As String
    strDigits = "98765432987654329876543298765432987654329876543298"
    Dim strSpecialCharacters As String
    strSpecialCharacters = ":;?=+)(*&^%$#@!:;?=+)(*&^%$#@!:;?=+)(*&^%$#@!:;?=+"

    'Use a randomly generated integer to further randomize the character lists
    Dim sngRnd As Single
    sngRnd = Rnd()
    If Right(CStr(sngRnd), 1) >= 8 Then
        strUpperCase = "GFDCBZXWVTSRQPNMLKJHGFDCBZXWVTSRQPNMLKJHGFDCBZXWVT"
        strLowerCase = "fdcbzxwvtsrqpnmkjhgfdcbzxwvtsrqpnmkjhgfdcbzxwvtsrq"
        strDigits = "76543298765432987654329876543298765432987654329876"
        strSpecialCharacters = "=+)(*&^%$#@!:;?=+)(*&^%$#@!:;?=+)(*&^%$#@!:;?=+()*"
    ElseIf Right(CStr(sngRnd), 1) > 5 Then
        strUpperCase = "CBZXWVTSRQPNMLKJHGFDCBZXWVTSRQPNMLKJHGFDCBZXWVTSRQ"
        strLowerCase = "bzxwvtsrqpnmkjhgfdcbzxwvtsrqpnmkjhgfdcbzxwvtsrqpnm"
        strDigits = "54329876543298765432987654329876543298765432987654"
        strSpecialCharacters = ")(*&^%$#@!:;?=+)(*&^%$#@!:;?=+)(*&^%$#@!:;?=+()*=+"
    ElseIf Right(CStr(sngRnd), 1) > 3 Then
        strUpperCase = "XWVTSRQPNMLKJHGFDCBZXWVTSRQPNMLKJHGFDCBZXWVTSRQPNM"
        strLowerCase = "wvtsrqpnmkjhgfdcbzxwvtsrqpnmkjhgfdcbzxwvtsrqpnmkjh"
        strDigits = "32987654329876543298765432987654329876543298765432"
        strSpecialCharacters = "^%$#@!:;?=+)(*&^%$#@!:;?=+)(*&^%$#@!:;?=+()*=+()*&"
    End If
    
    'Loop through each character in the password
    For i = 1 To Len(strPwd) - 1
        'Now we take the character we are looking at with "i" and then
        'we iterate through all the characters to the right.
        For i2 = i + 1 To Len(strPwd)
            'Check to see if the iterated character is the same as the "i"
            'value target character.
            If Mid(strPwd, i2, 1) = Mid(strPwd, i, 1) Then
                'Now we have to edit strPwd but do it so as we don't
                'change the length. That means we take all the parts
                'before i2, add the replacement characters, then add the part after i2.
                '
                'String before the character being replaced
                strNewPwd = Mid(strPwd, 1, i2 - 1)
                'Replacement character
                If funcIsCharType(Mid(strPwd, i2, 1)) = "u" Then
                    incrementUpper = incrementUpper + 1
                    strNewPwd = strNewPwd & Mid(strUpperCase, InStr(strUpperCase, Mid(strPwd, i2, 1)) + incrementUpper + 1, 1)
                ElseIf funcIsCharType(Mid(strPwd, i2, 1)) = "l" Then
                    incrementLower = incrementLower + 1
                    strNewPwd = strNewPwd & Mid(strLowerCase, InStr(strLowerCase, Mid(strPwd, i2, 1)) + incrementLower + 1, 1)
                ElseIf funcIsCharType(Mid(strPwd, i2, 1)) = "d" Then
                    incrementDigit = incrementDigit = 0 + 1
                    strNewPwd = strNewPwd & Mid(strDigits, InStr(strDigits, Mid(strPwd, i2, 1)) + incrementDigit + 1, 1)
                Else
                    incrementSpecial = incrementSpecial + 1
                    strNewPwd = strNewPwd & Mid(strSpecialCharacters, InStr(strSpecialCharacters, Mid(strPwd, i2, 1)) + incrementSpecial + 1, 1)
                End If
                'String after the character being replaced
                strNewPwd = strNewPwd & Mid(strPwd, i2 + 1, Len(strPwd))
                'Now assign the fully concatenated new password to the original variable
                strPwd = strNewPwd
            End If
        Next i2
    Next i
    funcConformPwdToRules_Duplicates = strPwd
End Function

Public Function funcConformPwdToRules_Minimums(ByVal strPwd As String) As String
    'These are the input cells on the Generator worksheet:
    'Number of Characters E5
    '
    'Minimum Lower Case E6
    'Minimum Upper Case E7
    'Minimum Digits E8
    'Minimum Special Characters E9
    funcConformPwdToRules_Minimums = strPwd
    
    'Declare variables
    Dim strCharacterList As String
    Dim strCharacterAdditional
    Dim i As Long

    Dim intCountUpper As Integer
    Dim intCountLower As Integer
    Dim intCountDigit As Integer
    Dim intCountSpecial As Integer

    If strPwd <> "" Then
        'UPPER CASE
        intCountUpper = funcCountCharacterByType(strPwd, "u")
        intCountLower = funcCountCharacterByType(strPwd, "l")
        intCountDigit = funcCountCharacterByType(strPwd, "d")
        intCountSpecial = funcCountCharacterByType(strPwd, "s")
        If intCountUpper < ThisWorkbook.Worksheets("Generator").Range("E7").Value Then
            Debug.Print "Short on UPPER"
            For i = 1 To Len(strPwd)
                If funcIsUpperCase(Mid(strPwd, i, 1)) = True Then
                    strCharacterList = strCharacterList & Mid(strPwd, i, 1)
                End If
            Next i
            strCharacterAdditional = funcGetAdditionalCharactersByType(strCharacterList, ThisWorkbook.Worksheets("Generator").Range("E7").Value - intCountUpper, "u")
            Do While Len(strCharacterAdditional) > 0
                For i = 2 To Len(strPwd) Step 2
                    If funcIsUpperCase(Mid(strPwd, i - 1, 1)) = False And funcIsUpperCase(Mid(strPwd, i, 1)) = False Then
                        strPwd = Mid(strPwd, 1, i - 1) & Mid(strCharacterAdditional, 1, 1) & Mid(strPwd, i, Len(strPwd))
                        If Len(strCharacterAdditional) > 1 Then
                            strCharacterAdditional = Mid(strCharacterAdditional, 2, Len(strCharacterAdditional))
                        Else
                            strCharacterAdditional = ""
                        End If
                    End If
                Next i
            Loop
        End If
        
        'LOWER CASE
        intCountUpper = funcCountCharacterByType(strPwd, "u")
        intCountLower = funcCountCharacterByType(strPwd, "l")
        intCountDigit = funcCountCharacterByType(strPwd, "d")
        intCountSpecial = funcCountCharacterByType(strPwd, "s")
        If intCountLower < ThisWorkbook.Worksheets("Generator").Range("E6").Value Then
            Debug.Print "Short on lower"
            For i = 1 To Len(strPwd)
                If funcIsLowerCase(Mid(strPwd, i, 1)) = True Then
                    strCharacterList = strCharacterList & Mid(strPwd, i, 1)
                End If
            Next i
            strCharacterAdditional = funcGetAdditionalCharactersByType(strCharacterList, ThisWorkbook.Worksheets("Generator").Range("E6").Value - intCountLower, "l")
            Do While Len(strCharacterAdditional) > 0
                For i = 2 To Len(strPwd) Step 2
                    If funcIsLowerCase(Mid(strPwd, i - 1, 1)) = False And funcIsLowerCase(Mid(strPwd, i, 1)) = False Then
                        strPwd = Mid(strPwd, 1, i - 1) & Mid(strCharacterAdditional, 1, 1) & Mid(strPwd, i, Len(strPwd))
                        If Len(strCharacterAdditional) > 1 Then
                            strCharacterAdditional = Mid(strCharacterAdditional, 2, Len(strCharacterAdditional))
                        Else
                            strCharacterAdditional = ""
                        End If
                    End If
                Next i
            Loop
        End If
        
        'DIGITS
        intCountUpper = funcCountCharacterByType(strPwd, "u")
        intCountLower = funcCountCharacterByType(strPwd, "l")
        intCountDigit = funcCountCharacterByType(strPwd, "d")
        intCountSpecial = funcCountCharacterByType(strPwd, "s")
        If intCountDigit < ThisWorkbook.Worksheets("Generator").Range("E8").Value Then
            Debug.Print "Short on digits"
            For i = 1 To Len(strPwd)
                If funcIsDigit(Mid(strPwd, i, 1)) = True Then
                    strCharacterList = strCharacterList & Mid(strPwd, i, 1)
                End If
            Next i
            strCharacterAdditional = funcGetAdditionalCharactersByType(strCharacterList, ThisWorkbook.Worksheets("Generator").Range("E8").Value - intCountDigit, "d")
            Do While Len(strCharacterAdditional) > 0
                For i = 2 To Len(strPwd) Step 2
                    If funcIsDigit(Mid(strPwd, i - 1, 1)) = False And funcIsDigit(Mid(strPwd, i, 1)) = False Then
                        strPwd = Mid(strPwd, 1, i - 1) & Mid(strCharacterAdditional, 1, 1) & Mid(strPwd, i, Len(strPwd))
                        If Len(strCharacterAdditional) > 1 Then
                            strCharacterAdditional = Mid(strCharacterAdditional, 2, Len(strCharacterAdditional))
                        Else
                            strCharacterAdditional = ""
                        End If
                    End If
                Next i
            Loop
        End If
        
        'SPECIAL CHARACTERS
        intCountUpper = funcCountCharacterByType(strPwd, "u")
        intCountLower = funcCountCharacterByType(strPwd, "l")
        intCountDigit = funcCountCharacterByType(strPwd, "d")
        intCountSpecial = funcCountCharacterByType(strPwd, "s")
        If intCountSpecial < ThisWorkbook.Worksheets("Generator").Range("E9").Value Then
            Debug.Print "Short on special"
            For i = 1 To Len(strPwd)
                If funcIsSpecialCharacter(Mid(strPwd, i, 1)) = True Then
                    strCharacterList = strCharacterList & Mid(strPwd, i, 1)
                End If
            Next i
            strCharacterAdditional = funcGetAdditionalCharactersByType(strCharacterList, ThisWorkbook.Worksheets("Generator").Range("E9").Value - intCountSpecial, "s")
            Do While Len(strCharacterAdditional) > 0
                For i = 2 To Len(strPwd) Step 2
                    If funcIsSpecialCharacter(Mid(strPwd, i - 1, 1)) = False And funcIsSpecialCharacter(Mid(strPwd, i, 1)) = False Then
                        strPwd = Mid(strPwd, 1, i - 1) & Mid(strCharacterAdditional, 1, 1) & Mid(strPwd, i, Len(strPwd))
                        If Len(strCharacterAdditional) > 1 Then
                            strCharacterAdditional = Mid(strCharacterAdditional, 2, Len(strCharacterAdditional))
                        Else
                            strCharacterAdditional = ""
                        End If
                    End If
                Next i
            Loop
        End If



'        Debug.Print "Total upper case: " & funcCountCharacterByType(strPwd, "u")
'        Debug.Print "Total lower case: " & funcCountCharacterByType(strPwd, "l")
'        Debug.Print "Total digits: " & funcCountCharacterByType(strPwd, "d")
'        Debug.Print "Total special: "; funcCountCharacterByType(strPwd, "s")
'        Debug.Print "checksum: " & (funcCountCharacterByType(strPwd, "u") + funcCountCharacterByType(strPwd, "l") + funcCountCharacterByType(strPwd, "d") + funcCountCharacterByType(strPwd, "s"))
        Debug.Print "Length: " & Len(strPwd)
        Debug.Print strPwd

        funcConformPwdToRules_Minimums = strPwd
    End If
End Function

Private Function funcGetAdditionalCharactersByType(ByVal strNotTheseList As String, ByVal intHowMany As Integer, ByVal strType As String) As String
    'Declare the character lists
    funcGetAdditionalCharactersByType = ""
    
    Dim strUpperCase As String
    strUpperCase = "KJHGFDCBZXWVTSRQPNMLKJHGFDCBZXWVTSRQPNMLKJHGFDCBZX"
    Dim strLowerCase As String
    strLowerCase = "jhgfdcbzxwvtsrqpnmkjhgfdcbzxwvtsrqpnmkjhgfdcbzxwvt"
    Dim strDigits As String
    strDigits = "98765432987654329876543298765432987654329876543298"
    Dim strSpecialCharacters As String
    strSpecialCharacters = ":;?=+)(*&^%$#@!:;?=+)(*&^%$#@!:;?=+)(*&^%$#@!:;?=+"

    If strNotTheseList <> "" Then
        Dim i As Integer
        Dim i2 As Integer
        For i = 1 To intHowMany
            If strType = "u" Then
                For i2 = 1 To Len(strUpperCase)
                    If Len(funcGetAdditionalCharactersByType) < intHowMany Then
                        If InStr(strNotTheseList, Mid(strUpperCase, i2, 1)) = 0 Then
                            funcGetAdditionalCharactersByType = funcGetAdditionalCharactersByType & Mid(strUpperCase, i2, 1)
                        End If
                    End If
                Next i2
            ElseIf strType = "l" Then
                For i2 = 1 To Len(strLowerCase)
                    If Len(funcGetAdditionalCharactersByType) < intHowMany Then
                        If InStr(strNotTheseList, Mid(strLowerCase, i2, 1)) = 0 Then
                            funcGetAdditionalCharactersByType = funcGetAdditionalCharactersByType & Mid(strLowerCase, i2, 1)
                        End If
                    End If
                Next i2
            ElseIf strType = "d" Then
                For i2 = 1 To Len(strDigits)
                    If Len(funcGetAdditionalCharactersByType) < intHowMany Then
                        If InStr(strNotTheseList, Mid(strDigits, i2, 1)) = 0 Then
                            funcGetAdditionalCharactersByType = funcGetAdditionalCharactersByType & Mid(strDigits, i2, 1)
                        End If
                    End If
                Next i2
            Else
                For i2 = 1 To Len(strSpecialCharacters)
                    If Len(funcGetAdditionalCharactersByType) < intHowMany Then
                        If InStr(strNotTheseList, Mid(strSpecialCharacters, i2, 1)) = 0 Then
                            funcGetAdditionalCharactersByType = funcGetAdditionalCharactersByType & Mid(strSpecialCharacters, i2, 1)
                        End If
                    End If
                Next i2
            End If
        Next i
    Else
        'There were none of this type in the password, now we just need to get that many from the list.
        If strType = "u" Then
            funcGetAdditionalCharactersByType = Mid(strUpperCase, 1, intHowMany)
        ElseIf strType = "l" Then
            funcGetAdditionalCharactersByType = Mid(strLowerCase, 1, intHowMany)
        ElseIf strType = "d" Then
            funcGetAdditionalCharactersByType = Mid(strDigits, 1, intHowMany)
        Else
            funcGetAdditionalCharactersByType = Mid(strSpecialCharacters, 1, intHowMany)
        End If
    End If
End Function

Public Function funcCountCharacterByType(ByVal strText As String, ByVal strType As String) As Integer
    funcCountCharacterByType = 0
    If strText <> "" Then
        Dim i As Integer
        For i = 1 To Len(strText)
            If UCase(strType) = "U" Then
                If funcIsCharType(Mid(strText, i, 1)) = "u" Then
                    funcCountCharacterByType = funcCountCharacterByType + 1
                End If
            ElseIf UCase(strType) = "L" Then
                If funcIsCharType(Mid(strText, i, 1)) = "l" Then
                    funcCountCharacterByType = funcCountCharacterByType + 1
                End If
            ElseIf UCase(strType) = "D" Then
                If funcIsCharType(Mid(strText, i, 1)) = "d" Then
                    funcCountCharacterByType = funcCountCharacterByType + 1
                End If
            Else
                If funcIsCharType(Mid(strText, i, 1)) = "s" Then
                    funcCountCharacterByType = funcCountCharacterByType + 1
                End If
            End If
        Next i
    End If
End Function



