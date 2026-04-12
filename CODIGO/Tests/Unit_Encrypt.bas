Attribute VB_Name = "Unit_Encrypt"
Option Explicit

' ==========================================================================
' Encrypt Test Suite
' Tests the RndCrypt symmetric XOR cipher: round-trip property for short
' ASCII strings, multi-word strings, empty strings, mixed character sets,
' and ciphertext divergence (output differs from plaintext).
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Sub test_suite_encrypt()
    Call UnitTesting.RunTest("encrypt_round_trip_short", test_encrypt_round_trip_short())
    Call UnitTesting.RunTest("encrypt_round_trip_multi", test_encrypt_round_trip_multi())
    Call UnitTesting.RunTest("encrypt_round_trip_empty", test_encrypt_round_trip_empty())
    Call UnitTesting.RunTest("encrypt_ciphertext_differs", test_encrypt_ciphertext_differs())
    Call UnitTesting.RunTest("encrypt_round_trip_mixed", test_encrypt_round_trip_mixed())
End Sub

Private Function test_encrypt_round_trip_short() As Boolean
    On Error GoTo Fail
    Dim original As String
    original = "Hello"
    Dim encrypted As String
    encrypted = RndCrypt(original, "K")
    Dim decrypted As String
    decrypted = RndCrypt(encrypted, "K")
    test_encrypt_round_trip_short = (decrypted = "Hello")
    Exit Function
Fail:
    test_encrypt_round_trip_short = False
End Function

Private Function test_encrypt_round_trip_multi() As Boolean
    On Error GoTo Fail
    Dim original As String
    original = "Hello World"
    Dim encrypted As String
    encrypted = RndCrypt(original, "Secret")
    Dim decrypted As String
    decrypted = RndCrypt(encrypted, "Secret")
    test_encrypt_round_trip_multi = (decrypted = "Hello World")
    Exit Function
Fail:
    test_encrypt_round_trip_multi = False
End Function

Private Function test_encrypt_round_trip_empty() As Boolean
    On Error GoTo Fail
    Dim original As String
    original = ""
    Dim encrypted As String
    encrypted = RndCrypt(original, "Key")
    Dim decrypted As String
    decrypted = RndCrypt(encrypted, "Key")
    test_encrypt_round_trip_empty = (decrypted = "")
    Exit Function
Fail:
    test_encrypt_round_trip_empty = False
End Function

Private Function test_encrypt_ciphertext_differs() As Boolean
    On Error GoTo Fail
    Dim original As String
    original = "Hello"
    Dim encrypted As String
    encrypted = RndCrypt(original, "Key")
    test_encrypt_ciphertext_differs = (encrypted <> "Hello")
    Exit Function
Fail:
    test_encrypt_ciphertext_differs = False
End Function

Private Function test_encrypt_round_trip_mixed() As Boolean
    On Error GoTo Fail
    Dim original As String
    original = "T3st!@#$"
    Dim encrypted As String
    encrypted = RndCrypt(original, "P@ss1")
    Dim decrypted As String
    decrypted = RndCrypt(encrypted, "P@ss1")
    test_encrypt_round_trip_mixed = (decrypted = "T3st!@#$")
    Exit Function
Fail:
    test_encrypt_round_trip_mixed = False
End Function

#End If
