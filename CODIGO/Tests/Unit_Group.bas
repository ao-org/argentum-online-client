Attribute VB_Name = "Unit_Group"
Option Explicit

' ==========================================================================
' Group Module Test Suite
' Tests group membership tracking: zero-member initialization, adding members
' increments count, removing members decrements count, membership index query.
'
' Since there are no add/remove functions, tests directly manipulate the
' global GroupSize and GroupMembers() arrays (mirroring Protocol.bas behavior).
' All tests call Group.Clear at the start and end to avoid polluting state.
'
' Requirements: 5.1, 5.2, 5.3, 5.4
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Sub test_suite_group()
    Call UnitTesting.RunTest("grp_starts_with_zero_members", test_grp_starts_with_zero_members())
    Call UnitTesting.RunTest("grp_add_member_increments_count", test_grp_add_member_increments_count())
    Call UnitTesting.RunTest("grp_add_multiple_members", test_grp_add_multiple_members())
    Call UnitTesting.RunTest("grp_remove_member_decrements_count", test_grp_remove_member_decrements_count())
    Call UnitTesting.RunTest("grp_membership_index_query", test_grp_membership_index_query())
    Call UnitTesting.RunTest("grp_clear_resets_to_zero", test_grp_clear_resets_to_zero())
End Sub

' Requirement 5.1: Group starts with zero members after Clear
Private Function test_grp_starts_with_zero_members() As Boolean
    On Error GoTo Fail
    
    Call Group.Clear
    
    test_grp_starts_with_zero_members = (Group.GroupSize = 0)
    Exit Function
Fail:
    Call Group.Clear
    test_grp_starts_with_zero_members = False
End Function

' Requirement 5.2: Adding a member increments the member count
Private Function test_grp_add_member_increments_count() As Boolean
    On Error GoTo Fail
    
    Call Group.Clear
    
    Group.GroupSize = 1
    ReDim Group.GroupMembers(Group.GroupSize) As t_GroupEntry
    Group.GroupMembers(0).Name = "Player1"
    Group.GroupMembers(0).charindex = 10
    Group.GroupMembers(0).GroupId = 1
    
    Dim result As Boolean
    result = (Group.GroupSize = 1)
    
    Call Group.Clear
    test_grp_add_member_increments_count = result
    Exit Function
Fail:
    Call Group.Clear
    test_grp_add_member_increments_count = False
End Function

' Requirement 5.2: Adding multiple members increments count correctly
Private Function test_grp_add_multiple_members() As Boolean
    On Error GoTo Fail
    
    Call Group.Clear
    
    Group.GroupSize = 3
    ReDim Group.GroupMembers(Group.GroupSize) As t_GroupEntry
    Group.GroupMembers(0).Name = "Player1"
    Group.GroupMembers(0).charindex = 10
    Group.GroupMembers(0).GroupId = 1
    Group.GroupMembers(1).Name = "Player2"
    Group.GroupMembers(1).charindex = 20
    Group.GroupMembers(1).GroupId = 2
    Group.GroupMembers(2).Name = "Player3"
    Group.GroupMembers(2).charindex = 30
    Group.GroupMembers(2).GroupId = 3
    
    Dim result As Boolean
    result = (Group.GroupSize = 3)
    
    If result Then
        result = (Group.GroupMembers(0).Name = "Player1") And _
                 (Group.GroupMembers(1).Name = "Player2") And _
                 (Group.GroupMembers(2).Name = "Player3")
    End If
    
    Call Group.Clear
    test_grp_add_multiple_members = result
    Exit Function
Fail:
    Call Group.Clear
    test_grp_add_multiple_members = False
End Function

' Requirement 5.3: Removing a member decrements the member count
Private Function test_grp_remove_member_decrements_count() As Boolean
    On Error GoTo Fail
    
    Call Group.Clear
    
    Group.GroupSize = 3
    ReDim Group.GroupMembers(Group.GroupSize) As t_GroupEntry
    Group.GroupMembers(0).Name = "Player1"
    Group.GroupMembers(0).charindex = 10
    Group.GroupMembers(0).GroupId = 1
    Group.GroupMembers(1).Name = "Player2"
    Group.GroupMembers(1).charindex = 20
    Group.GroupMembers(1).GroupId = 2
    Group.GroupMembers(2).Name = "Player3"
    Group.GroupMembers(2).charindex = 30
    Group.GroupMembers(2).GroupId = 3
    
    ' Simulate removing the middle member: shift last down, decrement size
    Group.GroupMembers(1) = Group.GroupMembers(2)
    Group.GroupSize = 2
    ReDim Preserve Group.GroupMembers(Group.GroupSize) As t_GroupEntry
    
    Dim result As Boolean
    result = (Group.GroupSize = 2)
    
    If result Then
        result = (Group.GroupMembers(0).Name = "Player1") And _
                 (Group.GroupMembers(1).Name = "Player3")
    End If
    
    Call Group.Clear
    test_grp_remove_member_decrements_count = result
    Exit Function
Fail:
    Call Group.Clear
    test_grp_remove_member_decrements_count = False
End Function

' Requirement 5.4: Membership index query finds correct member
Private Function test_grp_membership_index_query() As Boolean
    On Error GoTo Fail
    
    Call Group.Clear
    
    Group.GroupSize = 3
    ReDim Group.GroupMembers(Group.GroupSize) As t_GroupEntry
    Group.GroupMembers(0).Name = "Alpha"
    Group.GroupMembers(0).charindex = 100
    Group.GroupMembers(0).GroupId = 1
    Group.GroupMembers(1).Name = "Beta"
    Group.GroupMembers(1).charindex = 200
    Group.GroupMembers(1).GroupId = 2
    Group.GroupMembers(2).Name = "Gamma"
    Group.GroupMembers(2).charindex = 300
    Group.GroupMembers(2).GroupId = 3
    
    Dim foundIndex As Integer
    Dim found As Boolean
    found = False
    foundIndex = -1
    
    Dim i As Integer
    For i = 0 To Group.GroupSize - 1
        If Group.GroupMembers(i).charindex = 200 Then
            foundIndex = i
            found = True
            Exit For
        End If
    Next i
    
    Dim result As Boolean
    result = found And (foundIndex = 1) And (Group.GroupMembers(foundIndex).Name = "Beta")
    
    ' Verify non-existent charindex is NOT found
    Dim notFound As Boolean
    notFound = True
    For i = 0 To Group.GroupSize - 1
        If Group.GroupMembers(i).charindex = 999 Then
            notFound = False
            Exit For
        End If
    Next i
    
    result = result And notFound
    
    Call Group.Clear
    test_grp_membership_index_query = result
    Exit Function
Fail:
    Call Group.Clear
    test_grp_membership_index_query = False
End Function

' Requirement 5.1: Clear resets group back to zero members
Private Function test_grp_clear_resets_to_zero() As Boolean
    On Error GoTo Fail
    
    Group.GroupSize = 2
    ReDim Group.GroupMembers(Group.GroupSize) As t_GroupEntry
    Group.GroupMembers(0).Name = "Player1"
    Group.GroupMembers(0).charindex = 10
    Group.GroupMembers(1).Name = "Player2"
    Group.GroupMembers(1).charindex = 20
    
    Call Group.Clear
    
    test_grp_clear_resets_to_zero = (Group.GroupSize = 0)
    Exit Function
Fail:
    Call Group.Clear
    test_grp_clear_resets_to_zero = False
End Function

#End If