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
    VecLength = Sqr(vec.x * vec.x + vec.y * vec.y)
End Function

Function VecSqLength(ByRef vec As Vector2) As Single
    VecSqLength = vec.x * vec.x + vec.y * vec.y
End Function

Public Sub Normalize(ByRef vec As Vector2)
    Dim leng As Single
    leng = VecLength(vec)
    vec.x = vec.x / leng
    vec.y = vec.y / leng
End Sub

Function VAdd(a As Vector2, b As Vector2) As Vector2
    Dim ret As Vector2
    ret.x = a.x + b.x
    ret.y = a.y + b.y
    VAdd = ret
End Function

Function VSubs(a As Vector2, b As Vector2) As Vector2
    Dim ret As Vector2
    ret.x = a.x - b.x
    ret.y = a.y - b.y
    VSubs = ret
End Function

Function VMul(a As Vector2, b As Single) As Vector2
    Dim ret As Vector2
    ret.x = a.x * b
    ret.y = a.y * b
    VMul = ret
End Function

Public Function GetAngle(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
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
End Function

Public Function RadToDeg(radians As Double) As Double
    RadToDeg = radians * 180 / PI
End Function

Public Function FixAngle(ByVal angle As Single) As Single
    angle = angle Mod 360
    If angle < 0 Then
        angle = 360 + angle
    End If
    FixAngle = angle
End Function
