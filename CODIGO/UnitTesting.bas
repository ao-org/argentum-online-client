Attribute VB_Name = "UnitTesting"
Option Explicit

#If UNIT_TEST = 1 Then

Private TotalTests      As Long
Private PassedTests     As Long
Private FailedTests     As Long
Private FailedTestNames() As String
Private FailedTestCount As Long
Private TotalElapsed    As Double
Private sw              As clsInstrument

Private Const SUITE_COUNT As Long = 26

Public Sub Init()
    TotalTests = 0
    PassedTests = 0
    FailedTests = 0
    FailedTestCount = 0
    TotalElapsed = 0
    ReDim FailedTestNames(0 To 99)
    Set sw = New clsInstrument
    sw.start
End Sub

Public Sub RunTest(ByVal TestName As String, ByVal Result As Boolean)
    TotalTests = TotalTests + 1
    If Result Then
        PassedTests = PassedTests + 1
    Else
        FailedTests = FailedTests + 1
        FailedTestNames(FailedTestCount) = TestName
        FailedTestCount = FailedTestCount + 1
    End If
End Sub

Public Sub RunTestError(ByVal TestName As String, ByVal ErrorDesc As String)
    TotalTests = TotalTests + 1
    FailedTests = FailedTests + 1
    FailedTestNames(FailedTestCount) = TestName & ": " & ErrorDesc
    FailedTestCount = FailedTestCount + 1
End Sub

Public Function test_suite() As Boolean
    Dim i As Long
    For i = 1 To SUITE_COUNT
        Select Case i
            Case 1: Call Unit_Math.test_suite_math
            Case 2: Call Unit_Bitmask.test_suite_bitmask
            Case 3: Call Unit_Color.test_suite_color
            Case 4: Call Unit_ElapsedTime.test_suite_elapsed_time
            Case 5: Call Unit_Locale.test_suite_locale
            Case 6: Call Unit_MD5.test_suite_md5
            Case 7: Call Unit_MathExt.test_suite_math_ext
            Case 8: Call Unit_ArrayList.test_suite_arraylist
            Case 9: Call Unit_Encrypt.test_suite_encrypt
            Case 10: Call Unit_ValidNumber.test_suite_valid_number
            Case 11: Call Unit_QuickSort.test_suite_quick_sort
            Case 12: Call Unit_IniManager.test_suite_ini_manager
            Case 13: Call Unit_WorldTime.test_suite_world_time
            Case 14: Call Unit_JSON.test_suite_json
#If DIRECT_PLAY = 1 Then
            Case 15: Call Unit_NetRoundTrip.test_suite_net_round_trip
#End If
            Case 16: Call Unit_Cooldown.test_suite_cooldown
            Case 17: Call Unit_Group.test_suite_group
            Case 18: Call Unit_StringBuilderExt.test_suite_stringbuilder_ext
            Case 19: Call Unit_CryptoConvert.test_suite_crypto_convert
            Case 20: Call Unit_NumberFormat.test_suite_number_format
            Case 21: Call Unit_WordExtract.test_suite_word_extract
            Case 22: Call Unit_Instrument.test_suite_instrument
            Case 23: Call Unit_Language.test_suite_language
            Case 24: Call Unit_Settings.test_suite_settings
            Case 25: Call Unit_FieldParsing.test_suite_field_parsing
            Case 26: Call Unit_CharValidation.test_suite_char_validation
        End Select
    Next i
    test_suite = (FailedTests = 0)
End Function

Public Sub WriteResultsToFile(ByVal FilePath As String)
    On Error GoTo WriteResultsToFile_Err
    TotalElapsed = sw.ElapsedSeconds
    Dim f As Integer
    f = FreeFile
    Open FilePath For Output As #f
    Print #f, "=== AO20 CLIENT TEST REPORT ==="
    Print #f, "Total: " & TotalTests
    Print #f, "Passed: " & PassedTests
    Print #f, "Failed: " & FailedTests
    Print #f, "Elapsed: " & Format$(TotalElapsed, "0.000") & "s"
    
    If FailedTests > 0 Then
        Print #f, "Failed Tests:"
        Dim i As Long
        For i = 0 To FailedTestCount - 1
            Print #f, "  - " & FailedTestNames(i)
        Next i
    End If
    
    If FailedTests = 0 Then
        Print #f, "RESULT: PASS"
    Else
        Print #f, "RESULT: FAIL"
    End If
    Close #f
    Exit Sub
WriteResultsToFile_Err:
    Close #f
End Sub

#End If
