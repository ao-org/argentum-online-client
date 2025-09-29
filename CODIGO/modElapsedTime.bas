Attribute VB_Name = "modElapsedTime"
' === modTicksMasked.bas ===
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Const TICKS32 As Double = 4294967296#

' Legacy (keep for now, used by old code paths)
Public Function GetTickCount() As Long
    GetTickCount = timeGetTime() And &H7FFFFFFF
End Function

' New raw version (preferred)
Public Function GetTickCountRaw() As Long
    GetTickCountRaw = timeGetTime()
End Function

Public Function TicksElapsed(ByVal startTick As Long, ByVal currentTick As Long) As Double
    If currentTick >= startTick Then
        TicksElapsed = CDbl(currentTick - startTick)
    Else
        TicksElapsed = (TICKS32 - CDbl(startTick)) + CDbl(currentTick)
    End If
End Function

Public Function TickAfter(ByVal a As Long, ByVal b As Long) As Boolean
    TickAfter = (a - b) >= 0
End Function

Public Function PosMod(ByVal a As Double, ByVal m As Long) As Long
    If m <= 0 Then PosMod = 0: Exit Function
    Dim r As Double
    r = a - m * Fix(a / m)
    If r >= m Then r = r - m
    If r < 0 Then r = r + m
    PosMod = CLng(r)
End Function
