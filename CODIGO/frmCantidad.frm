VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2892
   ClientLeft      =   1632
   ClientTop       =   4416
   ClientWidth     =   4368
   ControlBox      =   0   'False
   FillColor       =   &H00C00000&
   ForeColor       =   &H8000000D&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   241
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   364
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1410
      MaxLength       =   9999
      TabIndex        =   0
      Top             =   1620
      Width           =   1560
   End
   Begin VB.Image cmdMas 
      Height          =   300
      Left            =   3120
      Tag             =   "0"
      Top             =   1605
      Width           =   300
   End
   Begin VB.Image cmdMenos 
      Height          =   300
      Left            =   960
      Tag             =   "0"
      Top             =   1605
      Width           =   300
   End
   Begin VB.Image cmdCerrar 
      Height          =   420
      Left            =   3900
      Tag             =   "0"
      Top             =   15
      Width           =   480
   End
   Begin VB.Image cmdTirarTodo 
      Height          =   420
      Left            =   2250
      Tag             =   "0"
      Top             =   2175
      Width           =   1680
   End
   Begin VB.Image cmdTirar 
      Height          =   420
      Left            =   435
      Tag             =   "0"
      Top             =   2160
      Width           =   1680
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
End
Attribute VB_Name = "frmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Public bmoving As Boolean

Public dX      As Integer

Public dy      As Integer
Option Explicit

' Constantes para SendMessage
Const WM_SYSCOMMAND As Long = &H112&

Const MOUSE_MOVE    As Long = &HF012&

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long


Private cBotonMas As clsGraphicalButton
Private cBotonMenos As clsGraphicalButton
Private cBotonTirarTodo As clsGraphicalButton
Private cBotonTirar As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton


Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call Aplicar_Transparencia(Me.hwnd, 240)
    
    'Call FormParser.Parse_Form(Me)
    Text1.SelStart = 1
    
    Me.Picture = LoadInterface("cantidad.bmp")
    
    Call LoadButtons
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCantidad.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub LoadButtons()

    Set cBotonTirarTodo = New clsGraphicalButton
    Set cBotonTirar = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonMas = New clsGraphicalButton
    Set cBotonMenos = New clsGraphicalButton


    Call cBotonTirarTodo.Initialize(cmdTirarTodo, "boton-tirar-todo-default.bmp", _
                                                "boton-tirar-todo-over.bmp", _
                                                "boton-tirar-todo-off.bmp", Me)
    
    Call cBotonTirar.Initialize(cmdTirar, "boton-tirar-default.bmp", _
                                                "boton-tirar-over.bmp", _
                                                "boton-tirar-off.bmp", Me)
                                                
    Call cBotonCerrar.Initialize(cmdCerrar, "boton-cerrar-default.bmp", _
                                                "boton-cerrar-over.bmp", _
                                                "boton-cerrar-off.bmp", Me)
                                                
    Call cBotonMas.Initialize(cmdMas, "boton-sm-mas-default.bmp", _
                                                "boton-sm-mas-over.bmp", _
                                                "boton-sm-mas-off.bmp", Me)
                                                
    Call cBotonMenos.Initialize(cmdMenos, "boton-sm-menos-default.bmp", _
                                                "boton-sm-menos-over.bmp", _
                                                "boton-sm-menos-off.bmp", Me)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCantidad.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub cmdCerrar_Click()
    
    On Error GoTo cmdCerrar_Click_Err
    
    Unload Me
    
    Exit Sub

cmdCerrar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCantidad.cmdCerrar_Click", Erl)
    Resume Next
    
End Sub


Private Sub cmdMas_Click()
    
    On Error GoTo cmdMas_Click_Err
    
    If Val(Text1.Text) < MAX_INVENTORY_OBJS Then
        Text1.Text = Val(Text1.Text) + 1
    End If
    
    Exit Sub

cmdMas_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCantidad.cmdMas_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdMenos_Click()
    
    On Error GoTo cmdMenos_Click_Err
    
    If Val(Text1.Text) > 0 Then
        Text1.Text = Val(Text1.Text) - 1
    End If
    
    Exit Sub

cmdMenos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCantidad.cmdMenos_Click", Erl)
    Resume Next
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoverForm Me.hwnd
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Text1_KeyPress_Err
    

    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0

        End If
    End If

    
    Exit Sub

Text1_KeyPress_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCantidad.Text1_KeyPress", Erl)
    Resume Next
    
End Sub


Private Sub Text1_Change()

    On Error GoTo errhandler

    If Val(Text1.Text) < 0 Then
        Text1.Text = "1"

    End If
    If frmMain.Inventario.SelectedItem <> FLAGORO Then
        If Val(Text1.Text) > MAX_INVENTORY_OBJS Then
            Text1.Text = "10000"
            Text1.SelStart = Len(Text1.Text)
    
        End If
    Else
        If Val(Text1.Text) > 100000 Then
            Text1.Text = "100000"
            Text1.SelStart = Len(Text1.Text)
        End If
    End If
    
    Exit Sub
errhandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    Text1.Text = "1"

End Sub


Private Sub cmdTirar_click()
    
    On Error GoTo tirar_click_Err
    
    
    If Not MainTimer.Check(TimersIndex.Drop) Then Exit Sub

    Call Sound.Sound_Play(SND_CLICK)

    If LenB(frmCantidad.Text1.Text) > 0 Then
        If Not IsNumeric(frmCantidad.Text1.Text) Then Exit Sub  'Should never happen
        
        If frmMain.Inventario.SelectedItem <> FLAGORO Then
        
            If ObjData(frmMain.Inventario.OBJIndex(frmMain.Inventario.SelectedItem)).Destruye = 0 Then
                Call WriteDrop(frmMain.Inventario.SelectedItem, frmCantidad.Text1.Text)
            Else
                PreguntaScreen = "El item se destruira al tirarlo ¿Esta seguro?"
                Pregunta = True
                DestItemSlot = frmMain.Inventario.SelectedItem
                DestItemCant = frmCantidad.Text1.Text
                
                PreguntaLocal = True
                PreguntaNUM = 1

            End If

        Else
            Call WriteDrop(frmMain.Inventario.SelectedItem, frmCantidad.Text1.Text)

        End If
        
        frmCantidad.Text1.Text = ""

    End If

    Unload Me

    
    Exit Sub

tirar_click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCantidad.tirar_click", Erl)
    Resume Next
    
End Sub



Private Sub cmdTirarTodo_click()
    
    On Error GoTo tirartodo_click_Err
    

    If Not MainTimer.Check(TimersIndex.Drop) Then Exit Sub

    Call Sound.Sound_Play(SND_CLICK)

    If frmMain.Inventario.SelectedItem = 0 Then Exit Sub

    If frmMain.Inventario.SelectedItem <> FLAGORO Then
        If ObjData(frmMain.Inventario.OBJIndex(frmMain.Inventario.SelectedItem)).Destruye = 0 Then
            Call WriteDrop(frmMain.Inventario.SelectedItem, frmMain.Inventario.Amount(frmMain.Inventario.SelectedItem))
        Else
            PreguntaScreen = "El item se destruira al tirarlo ¿Esta seguro?"
            Pregunta = True
            DestItemSlot = frmMain.Inventario.SelectedItem
            DestItemCant = frmMain.Inventario.Amount(frmMain.Inventario.SelectedItem)
            
            PreguntaLocal = True
            PreguntaNUM = 1

        End If

        Unload Me
    Else

        If UserGLD > 100000 Then
            Call WriteDrop(frmMain.Inventario.SelectedItem, 100000)
            Unload Me
        Else
            Call WriteDrop(frmMain.Inventario.SelectedItem, UserGLD)
            Unload Me

        End If

    End If

    frmCantidad.Text1.Text = ""

    
    Exit Sub

tirartodo_click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCantidad.tirartodo_click", Erl)
    Resume Next
    
End Sub
