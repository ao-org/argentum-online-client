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
    r As Byte
    A As Byte
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)

Sub Long_2_RGBA(Dest As RGBA, ByVal Src As Long)
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo Long_2_RGBA_Err
    
    Call CopyMemory(Dest, Src, 4)
    
    Exit Sub

Long_2_RGBA_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Color.Long_2_RGBA", Erl)
    Resume Next
    
End Sub

Function RGBA_2_Long(Color As RGBA) As Long
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo RGBA_2_Long_Err
    
    Call CopyMemory(RGBA_2_Long, Color, 4)
    
    Exit Function

RGBA_2_Long_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Color.RGBA_2_Long", Erl)
    Resume Next
    
End Function

Function RGBA_From_Long(ByVal Color As Long) As RGBA
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo RGBA_From_Long_Err
    
    Call CopyMemory(RGBA_From_Long, Color, 4)
    
    Exit Function

RGBA_From_Long_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Color.RGBA_From_Long", Erl)
    Resume Next
    
End Function

Function RGBA_From_Comp(ByVal r As Byte, ByVal G As Byte, ByVal B As Byte, Optional ByVal A As Byte = 255) As RGBA
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo RGBA_From_Comp_Err
    
    RGBA_From_Comp.r = r
    RGBA_From_Comp.G = G
    RGBA_From_Comp.B = B
    RGBA_From_Comp.A = A
    
    Exit Function

RGBA_From_Comp_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Color.RGBA_From_Comp", Erl)
    Resume Next
    
End Function

Function RGBA_From_vbColor(ByVal Color As Long) As RGBA
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo RGBA_From_Long_Err

    Call Long_2_RGBA(RGBA_From_vbColor, Color)

    RGBA_From_vbColor.A = RGBA_From_vbColor.r
    RGBA_From_vbColor.r = RGBA_From_vbColor.B
    RGBA_From_vbColor.B = RGBA_From_vbColor.A
    RGBA_From_vbColor.A = 255
    
    Exit Function

RGBA_From_Long_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Color.RGBA_From_Long", Erl)
    Resume Next
    
End Function

Sub SetRGBA(Color As RGBA, ByVal r As Byte, ByVal G As Byte, ByVal B As Byte, Optional ByVal A As Byte = 255)
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo SetRGBA_Err
    
    Color.r = r
    Color.G = G
    Color.B = B
    Color.A = A
    
    Exit Sub

SetRGBA_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Color.SetRGBA", Erl)
    Resume Next
    
End Sub

Sub Long_2_RGBAList(Dest() As RGBA, ByVal Src As Long)
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo Long_2_RGBAList_Err
    
    Dim i As Long
    
    For i = 0 To 3
        Call Long_2_RGBA(Dest(i), Src)
    Next
    
    Exit Sub

Long_2_RGBAList_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Color.Long_2_RGBAList", Erl)
    Resume Next
    
End Sub

Sub RGBAList(Dest() As RGBA, ByVal r As Byte, ByVal G As Byte, ByVal B As Byte, Optional ByVal A As Byte = 255)
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo RGBAList_Err
    
    Dim i As Long
    
    For i = 0 To 3
        Call SetRGBA(Dest(i), r, G, B, A)
    Next
    
    Exit Sub

RGBAList_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Color.RGBAList", Erl)
    Resume Next
    
End Sub


Sub RGBA_ToList(Dest() As RGBA, Color As RGBA)
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo RGBAList_Err
    
    Dim i As Long
    
    For i = 0 To 3
        Call SetRGBA(Dest(i), Color.r, Color.G, Color.B, Color.A)
    Next
    
    Exit Sub

RGBAList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.RGBA_ToList", Erl)
    Resume Next
    
End Sub

Sub Copy_RGBAList(Dest() As RGBA, Src() As RGBA)
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo Copy_RGBAList_Err
    
    Dim i As Long
    
    For i = 0 To 3
        Dest(i) = Src(i)
    Next
    
    Exit Sub

Copy_RGBAList_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Color.Copy_RGBAList", Erl)
    Resume Next
    
End Sub

Sub LerpRGBA(Dest As RGBA, A As RGBA, B As RGBA, ByVal Factor As Single)
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo LerpRGBA_Err
    
    Dim InvFactor As Single: InvFactor = (1 - Factor)

    Dest.r = A.r * InvFactor + B.r * Factor
    Dest.G = A.G * InvFactor + B.G * Factor
    Dest.B = A.B * InvFactor + B.B * Factor
    Dest.A = A.A * InvFactor + B.A * Factor
    
    Exit Sub

LerpRGBA_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Color.LerpRGBA", Erl)
    Resume Next
    
End Sub

Sub LerpRGB(Dest As RGBA, A As RGBA, B As RGBA, ByVal Factor As Single)
    '***************************************************
    'Author: Martín Trionfetti (HarThaoS)
    '***************************************************
    
    On Error GoTo LerpRGB_Err
    
    Dim InvFactor As Single: InvFactor = (1 - Factor)

    Dest.r = A.r * InvFactor + B.r * Factor
    Dest.G = A.G * InvFactor + B.G * Factor
    Dest.B = A.B * InvFactor + B.B * Factor
    
    Exit Sub

LerpRGB_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Color.LerpRGB", Erl)
    Resume Next
    
End Sub

Sub ModulateRGBA(Dest As RGBA, A As RGBA, B As RGBA)
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo ModulateRGBA_Err
    
    Dest.r = CLng(A.r) * B.r \ 255
    Dest.G = CLng(A.G) * B.G \ 255
    Dest.B = CLng(A.B) * B.B \ 255
    Dest.A = CLng(A.A) * B.A \ 255
    
    Exit Sub

ModulateRGBA_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Color.ModulateRGBA", Erl)
    Resume Next
    
End Sub

Sub AddRGBA(Dest As RGBA, A As RGBA, B As RGBA)
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo AddRGBA_Err
    
    Dest.r = min(CLng(A.r) + CLng(B.r), 255)
    Dest.G = min(CLng(A.G) + CLng(B.G), 255)
    Dest.B = min(CLng(A.B) + CLng(B.B), 255)
    Dest.A = min(CLng(A.A) + CLng(B.A), 255)
    
    Exit Sub

AddRGBA_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Color.AddRGBA", Erl)
    Resume Next
    
End Sub

Function vbColor_2_Long(Color As Long) As Long
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo vbColor_2_Long_Err
    
    Dim TmpColor As RGBA
    Call Long_2_RGBA(TmpColor, Color)

    TmpColor.A = TmpColor.r
    TmpColor.r = TmpColor.B
    TmpColor.B = TmpColor.A
    TmpColor.A = 255
    
    vbColor_2_Long = RGBA_2_Long(TmpColor)
    
    Exit Function

vbColor_2_Long_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Color.vbColor_2_Long", Erl)
    Resume Next
    
End Function

Sub Copy_RGBAList_WithAlpha(Dest() As RGBA, Src() As RGBA, ByVal Alpha As Byte)
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo Copy_RGBAList_WithAlpha_Err
    
    Dim i As Long
    
    For i = 0 To 3
        Dest(i) = Src(i)
        Dest(i).A = Alpha
    Next
    
    Exit Sub

Copy_RGBAList_WithAlpha_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Color.Copy_RGBAList_WithAlpha", Erl)
    Resume Next
    
End Sub

Function RGBA_ToString(Color As RGBA) As String
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo RGBA_ToString_Err
    
    RGBA_ToString = "RGBA(" & Color.r & ", " & Color.G & ", " & Color.B & ", " & Color.A & ")"
    
    Exit Function

RGBA_ToString_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Color.RGBA_ToString", Erl)
    Resume Next
    
End Function

