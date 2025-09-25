Attribute VB_Name = "ModBordes"
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
Option Explicit
Public COLOR_AZUL  As Long
Public COLOR_BORDE As Long
' funciones Api
'''''''''''''''''
' recupera el estilo del Listbox
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
' cambia el estilo
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
' refresca y vuelve a redibujar el control
Private Declare Function SetWindowPos _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal hWndInsertAfter As Long, _
                              ByVal x As Long, _
                              ByVal y As Long, _
                              ByVal cx As Long, _
                              ByVal cy As Long, _
                              ByVal wFlags As Long) As Long
' constantes para SetWindowPos
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
' para GetWindowLong - SetWindowLong
Private Const GWL_STYLE = (-16)
Private Const WS_BORDER = &H800000
Enum tShapeEstilo
    eCuadrado = 0
    eRedondeado = 4
End Enum

Public Sub Establecer_Borde(mObject As Object, _
                            frmParent As Form, _
                            Optional COLOR_BORDE As Long = vbBlack, _
                            Optional Border_style As BorderStyleConstants = vbBSDot, _
                            Optional Size_Border As Integer = 1, _
                            Optional estilo_Shape As tShapeEstilo = eCuadrado)
    On Error GoTo Establecer_Borde_Err
    Dim lng_Estilo As Long
    With mObject
        .Appearance = 0 ' flat
        lng_Estilo = GetWindowLong(.hWnd, GWL_STYLE)
        lng_Estilo = lng_Estilo And Not WS_BORDER ' sin borde
        ' aplica
        SetWindowLong .hWnd, GWL_STYLE, lng_Estilo
        ' refresh
        SetWindowPos .hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOOWNERZORDER Or SWP_NOSIZE Or SWP_NOZORDER
    End With
    Dim cTL             As Control
    Dim obj_Shape       As Shape
    Dim CustomShapeName As String
    CustomShapeName = "CustomBorderShape"
    ' crea un shape si no existe
    For Each cTL In frmParent.Controls
        If cTL.Name = CustomShapeName Then
            If TypeName(frmParent.Controls(CustomShapeName)) = "Shape" Then
                Set obj_Shape = cTL
                Exit For
            End If
        End If
    Next cTL
    If obj_Shape Is Nothing Then
        Set obj_Shape = frmParent.Controls.Add("vb.shape", CustomShapeName)
    End If
    With obj_Shape
        ' contenedor del shape
        Set .Container = mObject.Container
        ' posición
        If estilo_Shape = eCuadrado Then
            .Move mObject.Left - 30, mObject.Top - 30, mObject.Width + 60, mObject.Height + 60
        ElseIf estilo_Shape = eRedondeado Then
            .Move mObject.Left - 150, mObject.Top - 150, mObject.Width + 300, mObject.Height + 300
        Else
            Exit Sub
        End If
        'estilo de borde, color y tamaño
        .BorderStyle = Border_style
        If Border_style <> vbTransparent Then .BorderWidth = Size_Border
        .BorderColor = COLOR_BORDE
        .Shape = estilo_Shape
        .visible = True ' lo hace visible
        .ZOrder 0
    End With
    Set obj_Shape = Nothing
    Exit Sub
Establecer_Borde_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModBordes.Establecer_Borde", Erl)
    Resume Next
End Sub
