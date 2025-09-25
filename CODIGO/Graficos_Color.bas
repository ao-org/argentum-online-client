Attribute VB_Name = "Graficos_Color"
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
' ****************************************************
' Módulo de colores hecho por Alexis Caraballo (WyroX)
' Para una fácil conversión entre RGBA(4 bytes) y Long
' Nota: No uso D3DCOLORVALUE porque usa 4 singles
' ****************************************************
Option Explicit
Type RGBA
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Sub Long_2_RGBA(Dest As RGBA, ByVal src As Long)
    On Error GoTo Long_2_RGBA_Err
    Call CopyMemory(Dest, src, 4)
    Exit Sub
Long_2_RGBA_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.Long_2_RGBA", Erl)
    Resume Next
End Sub

Function RGBA_2_Long(color As RGBA) As Long
    On Error GoTo RGBA_2_Long_Err
    Call CopyMemory(RGBA_2_Long, color, 4)
    Exit Function
RGBA_2_Long_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.RGBA_2_Long", Erl)
    Resume Next
End Function

Function RGBA_From_Long(ByVal color As Long) As RGBA
    On Error GoTo RGBA_From_Long_Err
    Call CopyMemory(RGBA_From_Long, color, 4)
    Exit Function
RGBA_From_Long_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.RGBA_From_Long", Erl)
    Resume Next
End Function

Function RGBA_From_Comp(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, Optional ByVal A As Byte = 255) As RGBA
    On Error GoTo RGBA_From_Comp_Err
    RGBA_From_Comp.R = R
    RGBA_From_Comp.G = G
    RGBA_From_Comp.B = B
    RGBA_From_Comp.A = A
    Exit Function
RGBA_From_Comp_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.RGBA_From_Comp", Erl)
    Resume Next
End Function

Function RGBA_From_vbColor(ByVal color As Long) As RGBA
    On Error GoTo RGBA_From_Long_Err
    Call Long_2_RGBA(RGBA_From_vbColor, color)
    RGBA_From_vbColor.A = RGBA_From_vbColor.R
    RGBA_From_vbColor.R = RGBA_From_vbColor.B
    RGBA_From_vbColor.B = RGBA_From_vbColor.A
    RGBA_From_vbColor.A = 255
    Exit Function
RGBA_From_Long_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.RGBA_From_Long", Erl)
    Resume Next
End Function

Sub SetRGBA(color As RGBA, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, Optional ByVal A As Byte = 255)
    On Error GoTo SetRGBA_Err
    color.R = R
    color.G = G
    color.B = B
    color.A = A
    Exit Sub
SetRGBA_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.SetRGBA", Erl)
    Resume Next
End Sub

Sub Long_2_RGBAList(Dest() As RGBA, ByVal src As Long)
    On Error GoTo Long_2_RGBAList_Err
    Dim i As Long
    For i = 0 To 3
        Call Long_2_RGBA(Dest(i), src)
    Next
    Exit Sub
Long_2_RGBAList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.Long_2_RGBAList", Erl)
    Resume Next
End Sub

Sub RGBAList(Dest() As RGBA, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, Optional ByVal A As Byte = 255)
    On Error GoTo RGBAList_Err
    Dim i As Long
    For i = 0 To 3
        Call SetRGBA(Dest(i), R, G, B, A)
    Next
    Exit Sub
RGBAList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.RGBAList", Erl)
    Resume Next
End Sub

Sub RGBA_ToList(Dest() As RGBA, color As RGBA)
    On Error GoTo RGBAList_Err
    Dim i As Long
    For i = 0 To 3
        Call SetRGBA(Dest(i), color.R, color.G, color.B, color.A)
    Next
    Exit Sub
RGBAList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.RGBA_ToList", Erl)
    Resume Next
End Sub

Sub Copy_RGBAList(Dest() As RGBA, src() As RGBA)
    On Error GoTo Copy_RGBAList_Err
    Dim i As Long
    For i = 0 To 3
        Dest(i) = src(i)
    Next
    Exit Sub
Copy_RGBAList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.Copy_RGBAList", Erl)
    Resume Next
End Sub

Sub LerpRGBA(Dest As RGBA, A As RGBA, B As RGBA, ByVal Factor As Single)
    On Error GoTo LerpRGBA_Err
    Dim InvFactor As Single: InvFactor = (1 - Factor)
    Dest.R = A.R * InvFactor + B.R * Factor
    Dest.G = A.G * InvFactor + B.G * Factor
    Dest.B = A.B * InvFactor + B.B * Factor
    Dest.A = A.A * InvFactor + B.A * Factor
    Exit Sub
LerpRGBA_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.LerpRGBA", Erl)
    Resume Next
End Sub

Sub LerpRGB(Dest As RGBA, A As RGBA, B As RGBA, ByVal Factor As Single)
    On Error GoTo LerpRGB_Err
    Dim InvFactor As Single: InvFactor = (1 - Factor)
    Dest.R = A.R * InvFactor + B.R * Factor
    Dest.G = A.G * InvFactor + B.G * Factor
    Dest.B = A.B * InvFactor + B.B * Factor
    Exit Sub
LerpRGB_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.LerpRGB", Erl)
    Resume Next
End Sub

Sub ModulateRGBA(Dest As RGBA, A As RGBA, B As RGBA)
    On Error GoTo ModulateRGBA_Err
    Dest.R = CLng(A.R) * B.R \ 255
    Dest.G = CLng(A.G) * B.G \ 255
    Dest.B = CLng(A.B) * B.B \ 255
    Dest.A = CLng(A.A) * B.A \ 255
    Exit Sub
ModulateRGBA_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.ModulateRGBA", Erl)
    Resume Next
End Sub

Sub AddRGBA(Dest As RGBA, A As RGBA, B As RGBA)
    On Error GoTo AddRGBA_Err
    Dest.R = min(CLng(A.R) + CLng(B.R), 255)
    Dest.G = min(CLng(A.G) + CLng(B.G), 255)
    Dest.B = min(CLng(A.B) + CLng(B.B), 255)
    Dest.A = min(CLng(A.A) + CLng(B.A), 255)
    Exit Sub
AddRGBA_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.AddRGBA", Erl)
    Resume Next
End Sub

Function vbColor_2_Long(color As Long) As Long
    On Error GoTo vbColor_2_Long_Err
    Dim TmpColor As RGBA
    Call Long_2_RGBA(TmpColor, color)
    TmpColor.A = TmpColor.R
    TmpColor.R = TmpColor.B
    TmpColor.B = TmpColor.A
    TmpColor.A = 255
    vbColor_2_Long = RGBA_2_Long(TmpColor)
    Exit Function
vbColor_2_Long_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.vbColor_2_Long", Erl)
    Resume Next
End Function

Sub Copy_RGBAList_WithAlpha(Dest() As RGBA, src() As RGBA, ByVal alpha As Byte)
    On Error GoTo Copy_RGBAList_WithAlpha_Err
    Dim i As Long
    For i = 0 To 3
        Dest(i) = src(i)
        Dest(i).A = alpha
    Next
    Exit Sub
Copy_RGBAList_WithAlpha_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.Copy_RGBAList_WithAlpha", Erl)
    Resume Next
End Sub

Function RGBA_ToString(color As RGBA) As String
    On Error GoTo RGBA_ToString_Err
    RGBA_ToString = "RGBA(" & color.R & ", " & color.G & ", " & color.B & ", " & color.A & ")"
    Exit Function
RGBA_ToString_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.RGBA_ToString", Erl)
    Resume Next
End Function
