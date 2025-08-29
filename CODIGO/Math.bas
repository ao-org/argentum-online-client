Attribute VB_Name = "Math"
'    Argentum 20 - Game Client Program
'    Copyright (C) 2022 - Noland Studios
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'
Type Vector2
    x As Single
    y As Single
End Type

Function VecLength(ByRef vec As Vector2) As Single
    On Error Goto VecLength_Err
    VecLength = Sqr(vec.x * vec.x + vec.y * vec.y)
    Exit Function
VecLength_Err:
    Call TraceError(Err.Number, Err.Description, "Math.VecLength", Erl)
End Function

Function VecSqLength(ByRef vec As Vector2) As Single
    On Error Goto VecSqLength_Err
    VecSqLength = vec.x * vec.x + vec.y * vec.y
    Exit Function
VecSqLength_Err:
    Call TraceError(Err.Number, Err.Description, "Math.VecSqLength", Erl)
End Function

Public Sub Normalize(ByRef vec As Vector2)
    On Error Goto Normalize_Err
    Dim leng As Single
    leng = VecLength(vec)
    vec.x = vec.x / leng
    vec.y = vec.y / leng
    Exit Sub
Normalize_Err:
    Call TraceError(Err.Number, Err.Description, "Math.Normalize", Erl)
End Sub

Function VAdd(a As Vector2, b As Vector2) As Vector2
    On Error Goto VAdd_Err
    Dim ret As Vector2
    ret.x = a.x + b.x
    ret.y = a.y + b.y
    VAdd = ret
    Exit Function
VAdd_Err:
    Call TraceError(Err.Number, Err.Description, "Math.VAdd", Erl)
End Function

Function VSubs(a As Vector2, b As Vector2) As Vector2
    On Error Goto VSubs_Err
    Dim ret As Vector2
    ret.x = a.x - b.x
    ret.y = a.y - b.y
    VSubs = ret
    Exit Function
VSubs_Err:
    Call TraceError(Err.Number, Err.Description, "Math.VSubs", Erl)
End Function

Function VMul(a As Vector2, b As Single) As Vector2
    On Error Goto VMul_Err
    Dim ret As Vector2
    ret.x = a.x * b
    ret.y = a.y * b
    VMul = ret
    Exit Function
VMul_Err:
    Call TraceError(Err.Number, Err.Description, "Math.VMul", Erl)
End Function

Public Function GetAngle(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
    On Error Goto GetAngle_Err
    Dim XDiff As Double
    Dim YDiff As Double
    Dim TempAngle As Double

    YDiff = Abs(y2 - y1)

    If x1 = x2 And y1 = y2 Then Exit Function

    If YDiff = 0 And x1 < x2 Then
        GetAngle = 0
        Exit Function
    ElseIf YDiff = 0 And x1 > x2 Then
        GetAngle = 3.14159265358979
        Exit Function
    End If

    XDiff = Abs(x2 - x1)

    TempAngle = Atn(XDiff / YDiff)

    If y2 > y1 Then TempAngle = 3.14159265358979 - TempAngle
    If x2 < x1 Then TempAngle = -TempAngle
    TempAngle = 1.5707963267949 - TempAngle
    If TempAngle < 0 Then TempAngle = 6.28318530717959 + TempAngle

    GetAngle = TempAngle
    Exit Function
GetAngle_Err:
    Call TraceError(Err.Number, Err.Description, "Math.GetAngle", Erl)
End Function

Public Function RadToDeg(radians As Double) As Double
    On Error Goto RadToDeg_Err
    RadToDeg = radians * 180 / PI
    Exit Function
RadToDeg_Err:
    Call TraceError(Err.Number, Err.Description, "Math.RadToDeg", Erl)
End Function

Public Function FixAngle(ByVal angle As Single) As Single
    On Error Goto FixAngle_Err
    angle = angle Mod 360
    If angle < 0 Then
        angle = 360 + angle
    End If
    FixAngle = angle
    Exit Function
FixAngle_Err:
    Call TraceError(Err.Number, Err.Description, "Math.FixAngle", Erl)
End Function

Public Function Interpolate(ByVal a As Integer, ByVal b As Integer, ByVal t As Double) As Integer
    On Error Goto Interpolate_Err
    Interpolate = a + CInt((b - a) * t)
    Exit Function
Interpolate_Err:
    Call TraceError(Err.Number, Err.Description, "Math.Interpolate", Erl)
End Function

Public Function PointIsInsideRect(ByVal x As Integer, ByVal y As Integer, ByRef Rect As Rect) As Boolean
    On Error Goto PointIsInsideRect_Err
    PointIsInsideRect = x >= Rect.Left And x <= Rect.Right And y >= Rect.Top And y <= Rect.Bottom
    Exit Function
PointIsInsideRect_Err:
    Call TraceError(Err.Number, Err.Description, "Math.PointIsInsideRect", Erl)
End Function

Public Function OverlapRect(ByRef TargetRect As RECT, ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Heigth As Integer)
    On Error Goto OverlapRect_Err
    OverlapRect = True
    If PointIsInsideRect(x, y, TargetRect) Then Exit Function
    If PointIsInsideRect(x + Width, y, TargetRect) Then Exit Function
    If PointIsInsideRect(x, y + Heigth, TargetRect) Then Exit Function
    If PointIsInsideRect(x + Width, y + Heigth, TargetRect) Then Exit Function
    OverlapRect = TargetRect.Left >= x And TargetRect.Left <= (x + Width) And TargetRect.Top >= y And TargetRect.Bottom <= y + Width
    Exit Function
OverlapRect_Err:
    Call TraceError(Err.Number, Err.Description, "Math.OverlapRect", Erl)
End Function

Public Sub SetMask(ByRef Mask As Long, ByVal Value As Long)
    On Error Goto SetMask_Err
    Mask = Mask Or Value
    Exit Sub
SetMask_Err:
    Call TraceError(Err.Number, Err.Description, "Math.SetMask", Erl)
End Sub

Public Function IsSet(ByVal Mask As Long, ByVal Value As Long) As Boolean
    On Error Goto IsSet_Err
    IsSet = (Mask And Value) > 0
    Exit Function
IsSet_Err:
    Call TraceError(Err.Number, Err.Description, "Math.IsSet", Erl)
End Function

Public Sub UnsetMask(ByRef Mask As Long, ByVal Value As Long)
    On Error Goto UnsetMask_Err
    Mask = Mask And Not Value
    Exit Sub
UnsetMask_Err:
    Call TraceError(Err.Number, Err.Description, "Math.UnsetMask", Erl)
End Sub

Public Sub ResetMask(ByRef Mask As Long)
    On Error Goto ResetMask_Err
    Mask = 0
    Exit Sub
ResetMask_Err:
    Call TraceError(Err.Number, Err.Description, "Math.ResetMask", Erl)
End Sub

