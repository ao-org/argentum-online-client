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

Private Const SUITE_COUNT As Long = 15

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
            Case 1: Unit_Math.test_suite_math
            Case 2: Unit_Bitmask.test_suite_bitmask
            Case 3: Unit_Color.test_suite_color
            Case 4: Unit_ElapsedTime.test_suite_elapsed_time
            Case 5: Unit_Locale.test_suite_locale
            Case 6: Unit_MD5.test_suite_md5
            Case 7: Unit_MathExt.test_suite_math_ext
            Case 8: Unit_ArrayList.test_suite_arraylist
            Case 9: Unit_Encrypt.test_suite_encrypt
            Case 10: Unit_ValidNumber.test_suite_valid_number
            Case 11: Unit_QuickSort.test_suite_quick_sort
            Case 12: Unit_IniManager.test_suite_ini_manager
            Case 13: Unit_WorldTime.test_suite_world_time
            Case 14: Unit_FieldParsing.test_suite_field_parsing
            Case 15: Unit_CharValidation.test_suite_char_validation
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
