Attribute VB_Name = "Unit_ArrayList"
Option Explicit

' ==========================================================================
' ArrayList Test Suite
' Tests the clsArrayList byte array list: initialization, clearing,
' adding/removing items, existence checks, position lookups, value
' retrieval, last-item queries, and compaction after removal.
' ==========================================================================

#If UNIT_TEST = 1 Then

' Runs all ArrayList unit tests.
Public Sub test_suite_arraylist()
    Call UnitTesting.RunTest("arraylist_init_clear", test_init_clear())
    Call UnitTesting.RunTest("arraylist_add_exist", test_add_exist())
    Call UnitTesting.RunTest("arraylist_remove_exist", test_remove_exist())
    Call UnitTesting.RunTest("arraylist_item_pos", test_item_pos())
    Call UnitTesting.RunTest("arraylist_item_value", test_item_value())
    Call UnitTesting.RunTest("arraylist_get_last", test_get_last_item())
    Call UnitTesting.RunTest("arraylist_remove_compacts", test_remove_compacts())
End Sub

' Verifies after Initialize and Clear, all items are 0 and GetLastItem=0.
Private Function test_init_clear() As Boolean
    On Error GoTo Fail
    Dim al As New clsArrayList
    al.Initialize 1, 5
    al.Clear
    test_init_clear = (al.Item(1) = 0 And al.GetLastItem = 0)
    Exit Function
Fail:
    test_init_clear = False
End Function

' Verifies after Add(42), itemExist(42) returns True.
Private Function test_add_exist() As Boolean
    On Error GoTo Fail
    Dim al As New clsArrayList
    al.Initialize 1, 5
    al.Clear
    al.Add 42
    test_add_exist = (al.itemExist(42) = True)
    Exit Function
Fail:
    test_add_exist = False
End Function

' Verifies after Add(42) then Remove(42), itemExist(42) returns False.
Private Function test_remove_exist() As Boolean
    On Error GoTo Fail
    Dim al As New clsArrayList
    al.Initialize 1, 5
    al.Clear
    al.Add 42
    al.Remove 42
    test_remove_exist = (al.itemExist(42) = False)
    Exit Function
Fail:
    test_remove_exist = False
End Function

' Verifies after Add(42), itemPos(42) returns the correct index (1).
Private Function test_item_pos() As Boolean
    On Error GoTo Fail
    Dim al As New clsArrayList
    al.Initialize 1, 5
    al.Clear
    al.Add 42
    test_item_pos = (al.itemPos(42) = 1)
    Exit Function
Fail:
    test_item_pos = False
End Function

' Verifies after Add(42), Item at the correct index returns 42.
Private Function test_item_value() As Boolean
    On Error GoTo Fail
    Dim al As New clsArrayList
    al.Initialize 1, 5
    al.Clear
    al.Add 42
    Dim pos As Byte
    pos = al.itemPos(42)
    test_item_value = (al.Item(pos) = 42)
    Exit Function
Fail:
    test_item_value = False
End Function

' Verifies after Add(10) and Add(20), GetLastItem returns 20.
Private Function test_get_last_item() As Boolean
    On Error GoTo Fail
    Dim al As New clsArrayList
    al.Initialize 1, 5
    al.Clear
    al.Add 10
    al.Add 20
    test_get_last_item = (al.GetLastItem = 20)
    Exit Function
Fail:
    test_get_last_item = False
End Function

' Verifies after Add(10), Add(20), Remove(10), Item(1) = 20
' (items shift left to compact the list).
Private Function test_remove_compacts() As Boolean
    On Error GoTo Fail
    Dim al As New clsArrayList
    al.Initialize 1, 5
    al.Clear
    al.Add 10
    al.Add 20
    al.Remove 10
    test_remove_compacts = (al.Item(1) = 20)
    Exit Function
Fail:
    test_remove_compacts = False
End Function

#End If
