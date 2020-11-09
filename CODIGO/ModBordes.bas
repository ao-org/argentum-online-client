Attribute VB_Name = "ModBordes"
Option Explicit

Public COLOR_AZUL  As Long

Public COLOR_BORDE As Long

' funciones Api
'''''''''''''''''

' recupera el estilo del Listbox
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

' cambia el estilo
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' refresca y vuelve a redibujar el control
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

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

Public Sub Establecer_Borde(mObject As Object, frmParent As Form, Optional COLOR_BORDE As Long = vbBlack, Optional Border_style As BorderStyleConstants = vbBSDot, Optional Size_Border As Integer = 1, Optional estilo_Shape As tShapeEstilo = eCuadrado)

    Dim lng_Estilo As Long
    
    With mObject
        .Appearance = 0 ' flat
        lng_Estilo = GetWindowLong(.hwnd, GWL_STYLE)
        lng_Estilo = lng_Estilo And Not WS_BORDER ' sin borde
        
        ' aplica
        SetWindowLong .hwnd, GWL_STYLE, lng_Estilo
        
        ' refresh
        SetWindowPos .hwnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOOWNERZORDER Or SWP_NOSIZE Or SWP_NOZORDER

    End With
    
    Dim ctl             As Control

    Dim obj_Shape       As Shape
    
    Dim CustomShapeName As String

    CustomShapeName = "CustomBorderShape"

    ' crea un shape si no existe
    For Each ctl In frmParent.Controls

        If ctl.name = CustomShapeName Then
            If TypeName(frmParent.Controls(CustomShapeName)) = "Shape" Then
                Set obj_Shape = ctl
                Exit For

            End If

        End If

    Next ctl

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
        .Visible = True ' lo hace visible
        .ZOrder 0

    End With
    
    Set obj_Shape = Nothing

End Sub
