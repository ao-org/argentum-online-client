Attribute VB_Name = "Unit_Projectile"
Option Explicit

' ==========================================================================
' Projectile Direction Test Suite
' Tests ModPelota.CalcularTodo direction calculation and fps reset.
' Verifies that Pelota.DireccionX/DireccionY are computed correctly
' based on Cosa1/Cosa2 positions, and that Pelota.fps resets to 0.
'
' Requirements: 6.1, 6.2, 6.3, 6.4, 6.5, 6.6
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Sub test_suite_projectile()
    ' Example-based tests
    Call UnitTesting.RunTest("pelota_derecha", test_pelota_derecha())
    Call UnitTesting.RunTest("pelota_izquierda", test_pelota_izquierda())
    Call UnitTesting.RunTest("pelota_direccion_y", test_pelota_direccion_y())
    Call UnitTesting.RunTest("pelota_misma_posicion", test_pelota_misma_posicion())
    Call UnitTesting.RunTest("pelota_fps_reset", test_pelota_fps_reset())
    ' Property-based test
    Call UnitTesting.RunTest("pelota_pbt_fps_reset", test_pelota_pbt_fps_reset())
End Sub

' Requirement 6.1: Cosa1.X < Cosa2.X -> DireccionX > 0 (rightward)
' Requirement 6.6: Save/restore globals
Private Function test_pelota_derecha() As Boolean
    On Error GoTo Fail
    Dim origPelota As TPelota: origPelota = Pelota
    Dim origCosa1 As TCosa: origCosa1 = Cosa1
    Dim origCosa2 As TCosa: origCosa2 = Cosa2

    Cosa1.X = 10: Cosa1.Y = 50
    Cosa2.X = 100: Cosa2.Y = 50
    Call CalcularTodo

    test_pelota_derecha = (Pelota.DireccionX > 0)

    Pelota = origPelota: Cosa1 = origCosa1: Cosa2 = origCosa2
    Exit Function
Fail:
    Pelota = origPelota: Cosa1 = origCosa1: Cosa2 = origCosa2
    test_pelota_derecha = False
End Function

' Requirement 6.2: Cosa1.X > Cosa2.X -> DireccionX < 0 (leftward)
' Requirement 6.6: Save/restore globals
Private Function test_pelota_izquierda() As Boolean
    On Error GoTo Fail
    Dim origPelota As TPelota: origPelota = Pelota
    Dim origCosa1 As TCosa: origCosa1 = Cosa1
    Dim origCosa2 As TCosa: origCosa2 = Cosa2

    Cosa1.X = 100: Cosa1.Y = 50
    Cosa2.X = 10: Cosa2.Y = 50
    Call CalcularTodo

    test_pelota_izquierda = (Pelota.DireccionX < 0)

    Pelota = origPelota: Cosa1 = origCosa1: Cosa2 = origCosa2
    Exit Function
Fail:
    Pelota = origPelota: Cosa1 = origCosa1: Cosa2 = origCosa2
    test_pelota_izquierda = False
End Function

' Requirement 6.3: Cosa1.Y < Cosa2.Y -> DireccionY has correct sign
' Requirement 6.6: Save/restore globals
Private Function test_pelota_direccion_y() As Boolean
    On Error GoTo Fail
    Dim origPelota As TPelota: origPelota = Pelota
    Dim origCosa1 As TCosa: origCosa1 = Cosa1
    Dim origCosa2 As TCosa: origCosa2 = Cosa2

    ' Cosa1.Y < Cosa2.Y and Cosa1.X < Cosa2.X
    ' With the actual code logic, DireccionY stays positive when Cosa1.X < Cosa2.X
    Cosa1.X = 10: Cosa1.Y = 10
    Cosa2.X = 100: Cosa2.Y = 100
    Call CalcularTodo

    ' DireccionY should be positive (Cosa1.Y < Cosa2.Y and the ElseIf branch
    ' checks Cosa1.X < Cosa2.X which is true, keeping DireccionY positive)
    test_pelota_direccion_y = (Pelota.DireccionY > 0)

    Pelota = origPelota: Cosa1 = origCosa1: Cosa2 = origCosa2
    Exit Function
Fail:
    Pelota = origPelota: Cosa1 = origCosa1: Cosa2 = origCosa2
    test_pelota_direccion_y = False
End Function

' Requirement 6.4: Same position -> DireccionX=0, DireccionY=0
' Requirement 6.6: Save/restore globals
Private Function test_pelota_misma_posicion() As Boolean
    On Error GoTo Fail
    Dim origPelota As TPelota: origPelota = Pelota
    Dim origCosa1 As TCosa: origCosa1 = Cosa1
    Dim origCosa2 As TCosa: origCosa2 = Cosa2

    Cosa1.X = 50: Cosa1.Y = 50
    Cosa2.X = 50: Cosa2.Y = 50
    Call CalcularTodo

    test_pelota_misma_posicion = (Pelota.DireccionX = 0 And Pelota.DireccionY = 0)

    Pelota = origPelota: Cosa1 = origCosa1: Cosa2 = origCosa2
    Exit Function
Fail:
    Pelota = origPelota: Cosa1 = origCosa1: Cosa2 = origCosa2
    test_pelota_misma_posicion = False
End Function

' Requirement 6.5: fps reset to 0 after CalcularTodo
' Requirement 6.6: Save/restore globals
Private Function test_pelota_fps_reset() As Boolean
    On Error GoTo Fail
    Dim origPelota As TPelota: origPelota = Pelota
    Dim origCosa1 As TCosa: origCosa1 = Cosa1
    Dim origCosa2 As TCosa: origCosa2 = Cosa2

    ' Set fps to a non-zero value before calling CalcularTodo
    Pelota.fps = 10
    Cosa1.X = 20: Cosa1.Y = 30
    Cosa2.X = 80: Cosa2.Y = 70
    Call CalcularTodo

    test_pelota_fps_reset = (Pelota.fps = 0)

    Pelota = origPelota: Cosa1 = origCosa1: Cosa2 = origCosa2
    Exit Function
Fail:
    Pelota = origPelota: Cosa1 = origCosa1: Cosa2 = origCosa2
    test_pelota_fps_reset = False
End Function

' Feature: unit-test-coverage-tier4, Property 9: CalcularTodo fps reset invariant
' Validates: Requirements 6.5
Private Function test_pelota_pbt_fps_reset() As Boolean
    On Error GoTo Fail
    Dim origPelota As TPelota: origPelota = Pelota
    Dim origCosa1 As TCosa: origCosa1 = Cosa1
    Dim origCosa2 As TCosa: origCosa2 = Cosa2

    Dim i As Long
    For i = 1 To 110
        ' Generate varied positions using deterministic spread
        Cosa1.X = CInt((i * 7) Mod 200)
        Cosa1.Y = CInt((i * 13) Mod 200)
        Cosa2.X = CInt((i * 11) Mod 200)
        Cosa2.Y = CInt((i * 17) Mod 200)
        ' Set fps to non-zero before each call
        Pelota.fps = CInt(i Mod 100) + 1

        Call CalcularTodo

        If Pelota.fps <> 0 Then
            Pelota = origPelota: Cosa1 = origCosa1: Cosa2 = origCosa2
            test_pelota_pbt_fps_reset = False
            Exit Function
        End If
    Next i

    Pelota = origPelota: Cosa1 = origCosa1: Cosa2 = origCosa2
    test_pelota_pbt_fps_reset = True
    Exit Function
Fail:
    Pelota = origPelota: Cosa1 = origCosa1: Cosa2 = origCosa2
    test_pelota_pbt_fps_reset = False
End Function

#End If
