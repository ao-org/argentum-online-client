Attribute VB_Name = "ModBordes"
Option Explicit

Public COLOR_AZUL As Long
Public COLOR_BORDE As Long


' funciones Api
'''''''''''''''''

' recupera el estilo del Listbox
Private Declare Function GetWindowLong _
    Lib "user32" _
    Alias "GetWindowLongA" ( _
        ByVal hwnd As Long, _
        ByVal nIndex As Long) As Long

' cambia el estilo
Private Declare Function SetWindowLong _
    Lib "user32" _
    Alias "SetWindowLongA" ( _
        ByVal hwnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long

' refresca y vuelve a redibujar el control
Private Declare Function SetWindowPos _
    Lib "user32" ( _
        ByVal hwnd As Long, _
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


Public Sub Establecer_Borde( _
    mObject As Object, _
    frmParent As Form, _
    Optional COLOR_BORDE As Long = vbBlack, _
    Optional Border_style As BorderStyleConstants = vbBSDot, _
    Optional Size_Border As Integer = 1, _
    Optional estilo_Shape As tShapeEstilo = eCuadrado)

    Dim lng_Estilo As Long
    
    With mObject
        .Appearance = 0 ' flat
        lng_Estilo = GetWindowLong(.hwnd, GWL_STYLE)
        lng_Estilo = lng_Estilo And Not WS_BORDER ' sin borde
        
        ' aplica
        SetWindowLong .hwnd, GWL_STYLE, lng_Estilo
        
        ' refresh
        SetWindowPos .hwnd, 0, 0, 0, 0, 0, _
                        SWP_FRAMECHANGED Or _
                        SWP_NOACTIVATE Or _
                        SWP_NOMOVE Or _
                        SWP_NOOWNERZORDER Or _
                        SWP_NOSIZE Or SWP_NOZORDER
    End With
    
    Dim ctl As Control
    Dim obj_Shape As Shape
    
    Dim CustomShapeName As String
    CustomShapeName = "CustomBorderShape"

    ' crea un shape si no existe
    For Each ctl In frmParent.Controls
        If ctl.Name = CustomShapeName Then
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
