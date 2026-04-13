Attribute VB_Name = "Unit_CharValidation"
Option Explicit

' ==========================================================================
' Character and Input Validation Test Suite
' Tests the character validation, email validation, number formatting, and
' accent-removal functions in Mod_General: AsciiValidos,
' ValidDescriptionCharacters, LegalCharacter, CheckMailString, isValidEmail,
' BeautifyBigNumber, and Tilde. Covers acceptance/rejection boundaries,
' special-case characters, and format suffixes (Properties 2-6).
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Sub test_suite_char_validation()
    Call UnitTesting.RunTest("ascii_validos_accept_reject", test_ascii_validos_accept_reject())
    Call UnitTesting.RunTest("valid_desc_printable", test_valid_desc_printable())
    Call UnitTesting.RunTest("valid_desc_extended", test_valid_desc_extended())
    Call UnitTesting.RunTest("legal_char_allowed", test_legal_char_allowed())
    Call UnitTesting.RunTest("legal_char_forbidden", test_legal_char_forbidden())
    Call UnitTesting.RunTest("legal_char_backspace", test_legal_char_backspace())
    Call UnitTesting.RunTest("check_mail_valid_invalid", test_check_mail_valid_invalid())
    Call UnitTesting.RunTest("check_mail_no_dot", test_check_mail_no_dot())
    Call UnitTesting.RunTest("is_valid_email_cases", test_is_valid_email_cases())
    Call UnitTesting.RunTest("beautify_big_number", test_beautify_big_number())
    Call UnitTesting.RunTest("tilde_accents", test_tilde_accents())
End Sub

' Feature: unit-test-coverage-tier3, Property 2: AsciiValidos accepts only lowercase letters and spaces
' **Validates: Requirements 2.1, 2.2**
Private Function test_ascii_validos_accept_reject() As Boolean
    On Error GoTo Fail
    ' True for lowercase letters and spaces
    If Not AsciiValidos("hola mundo") Then
        test_ascii_validos_accept_reject = False
        Exit Function
    End If
    ' Note: AsciiValidos lowercases input first, so uppercase letters are accepted.
    ' False for string containing digits (digits are outside a-z after lowercasing)
    If AsciiValidos("abc123") Then
        test_ascii_validos_accept_reject = False
        Exit Function
    End If
    ' False for string containing punctuation
    If AsciiValidos("hola!") Then
        test_ascii_validos_accept_reject = False
        Exit Function
    End If
    ' True for empty string
    If Not AsciiValidos("") Then
        test_ascii_validos_accept_reject = False
        Exit Function
    End If
    test_ascii_validos_accept_reject = True
    Exit Function
Fail:
    test_ascii_validos_accept_reject = False
End Function


' Feature: unit-test-coverage-tier3, Property 3: ValidDescriptionCharacters accepts printable range
' **Validates: Requirements 2.3**
Private Function test_valid_desc_printable() As Boolean
    On Error GoTo Fail
    ' True for printable ASCII (codes 32-126)
    If Not ValidDescriptionCharacters("Hello World 123!") Then
        test_valid_desc_printable = False
        Exit Function
    End If
    ' False for control character Chr$(1)
    If ValidDescriptionCharacters("abc" & Chr$(1) & "def") Then
        test_valid_desc_printable = False
        Exit Function
    End If
    test_valid_desc_printable = True
    Exit Function
Fail:
    test_valid_desc_printable = False
End Function

' Feature: unit-test-coverage-tier3, Property 3: ValidDescriptionCharacters accepts printable range
' **Validates: Requirements 2.4**
Private Function test_valid_desc_extended() As Boolean
    On Error GoTo Fail
    ' True for string with extended Latin chars (code >= 160)
    Dim extended As String
    extended = "abc" & Chr$(160) & Chr$(200) & Chr$(255)
    test_valid_desc_extended = ValidDescriptionCharacters(extended)
    Exit Function
Fail:
    test_valid_desc_extended = False
End Function

' Feature: unit-test-coverage-tier3, Property 4: LegalCharacter allowed-set invariant
' **Validates: Requirements 2.5**
Private Function test_legal_char_allowed() As Boolean
    On Error GoTo Fail
    ' Space (32), letter A (65), digit 1 (49) should all be allowed
    test_legal_char_allowed = LegalCharacter(32) And _
                              LegalCharacter(65) And _
                              LegalCharacter(49)
    Exit Function
Fail:
    test_legal_char_allowed = False
End Function

' Feature: unit-test-coverage-tier3, Property 4: LegalCharacter allowed-set invariant
' **Validates: Requirements 2.5**
Private Function test_legal_char_forbidden() As Boolean
    On Error GoTo Fail
    ' Asterisk (42), forward slash (47), control char (10) should be forbidden
    If LegalCharacter(42) Then
        test_legal_char_forbidden = False
        Exit Function
    End If
    If LegalCharacter(47) Then
        test_legal_char_forbidden = False
        Exit Function
    End If
    If LegalCharacter(10) Then
        test_legal_char_forbidden = False
        Exit Function
    End If
    test_legal_char_forbidden = True
    Exit Function
Fail:
    test_legal_char_forbidden = False
End Function

' Feature: unit-test-coverage-tier3, Property 4: LegalCharacter allowed-set invariant
' **Validates: Requirements 2.6**
Private Function test_legal_char_backspace() As Boolean
    On Error GoTo Fail
    ' Backspace (8) is a special-case allowed character
    test_legal_char_backspace = LegalCharacter(8)
    Exit Function
Fail:
    test_legal_char_backspace = False
End Function

' **Validates: Requirements 2.7**
Private Function test_check_mail_valid_invalid() As Boolean
    On Error GoTo Fail
    ' True for well-formed email
    If Not CheckMailString("[email protected]") Then
        test_check_mail_valid_invalid = False
        Exit Function
    End If
    ' False for missing @ symbol
    If CheckMailString("userexample.com") Then
        test_check_mail_valid_invalid = False
        Exit Function
    End If
    test_check_mail_valid_invalid = True
    Exit Function
Fail:
    test_check_mail_valid_invalid = False
End Function

' **Validates: Requirements 2.8**
Private Function test_check_mail_no_dot() As Boolean
    On Error GoTo Fail
    ' False for email with no dot after @
    test_check_mail_no_dot = Not CheckMailString("user@domain")
    Exit Function
Fail:
    test_check_mail_no_dot = False
End Function

' **Validates: Requirements 2.9**
Private Function test_is_valid_email_cases() As Boolean
    On Error GoTo Fail
    ' True for well-formed email
    If Not isValidEmail("[email protected]") Then
        test_is_valid_email_cases = False
        Exit Function
    End If
    ' False for missing @
    If isValidEmail("userexample.com") Then
        test_is_valid_email_cases = False
        Exit Function
    End If
    ' False for trailing dot
    If isValidEmail("[email protected].") Then
        test_is_valid_email_cases = False
        Exit Function
    End If
    test_is_valid_email_cases = True
    Exit Function
Fail:
    test_is_valid_email_cases = False
End Function

' Feature: unit-test-coverage-tier3, Property 5: BeautifyBigNumber suffix matches magnitude
' **Validates: Requirements 2.10**
Private Function test_beautify_big_number() As Boolean
    On Error GoTo Fail
    ' Raw for <= 10000
    Dim raw As String
    raw = BeautifyBigNumber(5000)
    If raw <> "5000" Then
        test_beautify_big_number = False
        Exit Function
    End If
    ' Boundary: exactly 10000 returns raw
    Dim boundary As String
    boundary = BeautifyBigNumber(10000)
    If boundary <> "10000" Then
        test_beautify_big_number = False
        Exit Function
    End If
    ' "K" suffix for > 10000
    Dim kSuffix As String
    kSuffix = BeautifyBigNumber(50000)
    If Right$(kSuffix, 1) <> "K" Then
        test_beautify_big_number = False
        Exit Function
    End If
    ' "KK" suffix for > 10000000
    Dim kkSuffix As String
    kkSuffix = BeautifyBigNumber(50000000)
    If Right$(kkSuffix, 2) <> "KK" Then
        test_beautify_big_number = False
        Exit Function
    End If
    ' "KKK" suffix for > 1000000000
    Dim kkkSuffix As String
    kkkSuffix = BeautifyBigNumber(2000000000)
    If Right$(kkkSuffix, 3) <> "KKK" Then
        test_beautify_big_number = False
        Exit Function
    End If
    test_beautify_big_number = True
    Exit Function
Fail:
    test_beautify_big_number = False
End Function

' Feature: unit-test-coverage-tier3, Property 6: Tilde produces uppercase with no accented vowels
' **Validates: Requirements 2.11**
Private Function test_tilde_accents() As Boolean
    On Error GoTo Fail
    ' "café" should become "CAFE" (accent removed, uppercased)
    Dim result As String
    result = Tilde("caf" & Chr$(233))
    If result <> "CAFE" Then
        test_tilde_accents = False
        Exit Function
    End If
    ' Plain lowercase should just uppercase
    If Tilde("hello") <> "HELLO" Then
        test_tilde_accents = False
        Exit Function
    End If
    test_tilde_accents = True
    Exit Function
Fail:
    test_tilde_accents = False
End Function

#End If