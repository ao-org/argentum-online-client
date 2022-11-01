VERSION 5.00
Begin VB.Form frmComerciar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Comerciando con el NPC"
   ClientHeight    =   7140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   595
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrNumber 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox interface 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3660
      Left            =   600
      MousePointer    =   99  'Custom
      ScaleHeight     =   305
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   574
      TabIndex        =   1
      Top             =   1605
      Width           =   6885
   End
   Begin VB.TextBox cantidad 
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
      Height          =   210
      Left            =   3675
      TabIndex        =   0
      Text            =   "1"
      Top             =   6555
      Width           =   810
   End
   Begin VB.Image cmdCerrar 
      Height          =   375
      Left            =   7680
      Top             =   0
      Width           =   495
   End
   Begin VB.Image cmdMas 
      Height          =   315
      Left            =   4650
      Tag             =   "1"
      Top             =   6510
      Width           =   315
   End
   Begin VB.Image cmdMenos 
      Height          =   315
      Left            =   3195
      Tag             =   "1"
      Top             =   6525
      Width           =   315
   End
   Begin VB.Label lbldesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "descripción"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   5940
      Width           =   3135
   End
   Begin VB.Label lblnombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Vacío)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   5610
      Width           =   3135
   End
   Begin VB.Label lblcosto 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      TabIndex        =   2
      Top             =   5550
      Width           =   1215
   End
   Begin VB.Image cmdVender 
      Height          =   420
      Left            =   5505
      Tag             =   "0"
      Top             =   6465
      Width           =   1830
   End
   Begin VB.Image cmdComprar 
      Height          =   420
      Left            =   825
      Tag             =   "0"
      Top             =   6465
      Width           =   1830
   End
End
Attribute VB_Name = "frmComerciar"
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
Option Explicit

Const WM_SYSCOMMAND As Long = &H112&

Const MOUSE_MOVE    As Long = &HF012&

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Public LastIndex1           As Integer

Public LasActionBuy         As Boolean

Private m_Number            As Integer

Private m_Increment         As Integer

Private m_Interval          As Integer

' Declaro los inventarios acá para poder manejar los eventos de drop
Public WithEvents InvComUsu As clsGrapchicalInventory ' Inventario del usuario visible en el comercio
Attribute InvComUsu.VB_VarHelpID = -1

Public WithEvents InvComNpc As clsGrapchicalInventory ' Inventario con los items que ofrece el npc
Attribute InvComNpc.VB_VarHelpID = -1

Private cBotonComprar As clsGraphicalButton
Private cBotonVender As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton
Private cBotonMas As clsGraphicalButton
Private cBotonMenos As clsGraphicalButton

Private Sub MoverForm()
    
    On Error GoTo moverForm_Err
    

    Dim res As Long

    ReleaseCapture
    res = SendMessage(Me.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)

    
    Exit Sub

moverForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.moverForm", Erl)
    Resume Next
    
End Sub

Private Sub cmdcerrar_Click()
    Unload Me
End Sub

Private Sub cmdComprar_DblClick()
    cmdComprar_Click
End Sub

Private Sub cmdVender_DblClick()
     cmdVender_Click
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    cantidad.BackColor = RGB(18, 19, 13)

    Me.Picture = LoadInterface("comerciar.bmp")
    Call LoadButtons
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub LoadButtons()
       
    Set cBotonComprar = New clsGraphicalButton
    Set cBotonVender = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonMas = New clsGraphicalButton
    Set cBotonMenos = New clsGraphicalButton


    Call cBotonComprar.Initialize(cmdComprar, "boton-comprar-default.bmp", _
                                                "boton-comprar-over.bmp", _
                                                "boton-comprar-off.bmp", Me)
    
    Call cBotonVender.Initialize(cmdVender, "boton-vender-default.bmp", _
                                                "boton-vender-over.bmp", _
                                                "boton-vender-off.bmp", Me)
                                                
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
Private Sub cantidad_KeyPress(KeyAscii As Integer)
    
    On Error GoTo cantidad_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If

    If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0

        End If

    End If

    
    Exit Sub

cantidad_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.cantidad_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub cmdMas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_Increment = 1

    tmrNumber.Interval = 30
    tmrNumber.Enabled = True
End Sub

Private Sub cmdMenos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    cantidad.Text = str((Val(cantidad.Text) - 1))
    m_Increment = -1
    
    tmrNumber.Interval = 30
    tmrNumber.Enabled = True
End Sub


Private Sub cmdMenos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tmrNumber.Enabled = False
End Sub
Private Sub cmdMas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tmrNumber.Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub cmdComprar_Click()
    If Not IsNumeric(cantidad.Text) Then Exit Sub
    If Val(cantidad.Text) <= 0 Then Exit Sub
    
     If InvComNpc.SelectedItem <= 0 Then Exit Sub
 
    LasActionBuy = True

    If UserGLD >= InvComNpc.Valor(InvComNpc.SelectedItem) * Val(cantidad) Then
        Call WriteCommerceBuy(InvComNpc.SelectedItem, cantidad.Text)
    Else
        AddtoRichTextBox frmMain.RecTxt, "No tenés suficiente oro.", 2, 51, 223, 1, 1

    End If
End Sub

Private Sub cmdVender_Click()
    If Not IsNumeric(cantidad.Text) Then Exit Sub
    If Val(cantidad.Text) <= 0 Then Exit Sub
    
     If InvComUsu.SelectedItem <= 0 Then Exit Sub
            
    LasActionBuy = False
    
    Call WriteCommerceSell(InvComUsu.SelectedItem, min(Val(cantidad.Text), InvComUsu.Amount(InvComUsu.SelectedItem)))
End Sub




Private Sub addRemove_Click(Index As Integer)
    
    On Error GoTo addRemove_Click_Err
    
    Call Sound.Sound_Play(SND_CLICK)

    Select Case Index

        Case 0
            cantidad = cantidad - 1

        Case 1
            cantidad = cantidad + 1

    End Select

    
    Exit Sub

addRemove_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.addRemove_Click", Erl)
    Resume Next
    
End Sub

Private Sub cantidad_Change()
    
    On Error GoTo cantidad_Change_Err
    

    If Val(cantidad.Text) < 0 Then
        cantidad.Text = 1
        m_Number = 1
    ElseIf Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
        cantidad.Text = 1
        m_Number = 1
    Else
        m_Number = Val(cantidad.Text)

    End If
    
    cantidad.SelStart = Len(cantidad.Text)
    
    InvComUsu.ReDraw

    
    Exit Sub

cantidad_Change_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.cantidad_Change", Erl)
    Resume Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error GoTo Form_Unload_Err
    
    If frmComerciar.Visible Then
        Call WriteCommerceEnd
    End If
    
    Exit Sub

Form_Unload_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.Form_Unload", Erl)
    Resume Next
    
End Sub



Private Sub interface_Click()
    
    On Error GoTo interface_Click_Err
    
    
    If InvComNpc.ClickedInside Then
        ' Cliqueé en la tienda, deselecciono el inventario
        Call InvComUsu.SeleccionarItem(0)
        
    ElseIf InvComUsu.ClickedInside Then
        ' Cliqueé en el inventario, deselecciono la tienda
        Call InvComNpc.SeleccionarItem(0)

    End If

    
    Exit Sub

interface_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.interface_Click", Erl)
    Resume Next
    
End Sub

Private Sub interface_DblClick()
    
    On Error GoTo interface_DblClick_Err
    

    If InvComNpc.ClickedInside Then
        
        If Not InvComNpc.IsItemSelected Then Exit Sub
    
        LasActionBuy = True

      '  If UserGLD >= InvComNpc.Valor(InvComNpc.SelectedItem) * Val(cantidad) Then
      '      Call WriteCommerceBuy(InvComNpc.SelectedItem, cantidad.Text)
      '  Else
      '      AddtoRichTextBox frmMain.RecTxt, "No tenés suficiente oro.", 2, 51, 223, 1, 1

      '  End If
        
    ElseIf InvComUsu.ClickedInside Then
    
        If Not InvComUsu.IsItemSelected Then Exit Sub
    
        ' Hacemos acción del doble clic correspondiente
        Dim ObjType As Byte

        ObjType = ObjData(InvComUsu.OBJIndex(InvComUsu.SelectedItem)).ObjType
        
        If UserMeditar Then Exit Sub
        If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
        
        Select Case ObjType

            Case eObjType.otArmadura, eObjType.otESCUDO, eObjType.otmagicos, eObjType.otFlechas, eObjType.otCASCO, eObjType.otNudillos, eObjType.otAnillos
                Call WriteEquipItem(InvComUsu.SelectedItem)
                
            Case eObjType.otWeapon

                If ObjData(InvComUsu.OBJIndex(InvComUsu.SelectedItem)).proyectil = 1 And InvComUsu.Equipped(InvComUsu.SelectedItem) Then
                    Call WriteUseItem(InvComUsu.SelectedItem)
                Else
                    Call WriteEquipItem(InvComUsu.SelectedItem)

                End If
                
            Case eObjType.OtHerramientas

                If InvComUsu.Equipped(InvComUsu.SelectedItem) Then
                    Call WriteUseItem(InvComUsu.SelectedItem)
                Else
                    Call WriteEquipItem(InvComUsu.SelectedItem)

                End If
                 
            Case Else
                Call WriteUseItem(InvComUsu.SelectedItem)

        End Select

    End If

    
    Exit Sub

interface_DblClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.interface_DblClick", Erl)
    Resume Next
    
End Sub

Private Sub interface_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo interface_KeyDown_Err
    

    ' Referencia temporal al inventario que corresponda
    Dim CurrentInventory As clsGrapchicalInventory

    If InvComNpc.ClickedInside Then
        Set CurrentInventory = InvComNpc
    ElseIf InvComUsu.ClickedInside Then
        Set CurrentInventory = InvComUsu
    Else
        Exit Sub

    End If

    ' Procesamos las teclas para moverse por el inventario
    Select Case KeyCode
    
        Case vbKeyRight

            If CurrentInventory.SelectedItem < CurrentInventory.MaxSlots Then
                Call CurrentInventory.SeleccionarItem(CurrentInventory.SelectedItem + 1)

            End If
            
        Case vbKeyLeft

            If CurrentInventory.SelectedItem > 1 Then
                Call CurrentInventory.SeleccionarItem(CurrentInventory.SelectedItem - 1)

            End If
            
        Case vbKeyUp

            If CurrentInventory.SelectedItem > CurrentInventory.Columns Then
                Call CurrentInventory.SeleccionarItem(CurrentInventory.SelectedItem - CurrentInventory.Columns)

            End If
            
        Case vbKeyDown

            If CurrentInventory.SelectedItem < CurrentInventory.MaxSlots - CurrentInventory.Columns Then
                Call CurrentInventory.SeleccionarItem(CurrentInventory.SelectedItem + CurrentInventory.Columns)

            End If
    
    End Select
    
    ' Limpiamos
    Set CurrentInventory = Nothing

    
    Exit Sub

interface_KeyDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.interface_KeyDown", Erl)
    Resume Next
    
End Sub

Private Sub InvComUsu_ItemDropped(ByVal Drag As Integer, ByVal Drop As Integer, ByVal x As Integer, ByVal y As Integer)
    
    On Error GoTo InvComUsu_ItemDropped_Err
    

    ' Si soltó dentro del mismo inventario
    If Drop > 0 Then
        If Drag <> Drop Then
            ' Movemos el item dentro del inventario
            Call WriteItemMove(Drag, Drop)
        End If
    Else

        ' Si lo soltó dentro de la tienda
        If InvComNpc.GetSlot(x, y) > 0 Then
            ' Vendemos el item
            LasActionBuy = False
            Call WriteCommerceSell(Drag, max(Val(cantidad.Text), InvComUsu.Amount(InvComUsu.SelectedItem)))

        End If

    End If

    
    Exit Sub

InvComUsu_ItemDropped_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.InvComUsu_ItemDropped", Erl)
    Resume Next
    
End Sub

Private Sub InvComNpc_ItemDropped(ByVal Drag As Integer, ByVal Drop As Integer, ByVal x As Integer, ByVal y As Integer)
    
    On Error GoTo InvComNpc_ItemDropped_Err
    

    ' Si lo soltó dentro del inventario
    If InvComUsu.GetSlot(x, y) > 0 Then
        ' Compramos el item
        LasActionBuy = True

        ' Si tiene suficiente oro
        If UserGLD >= InvComNpc.Valor(Drag) * Val(cantidad.Text) Then
            Call WriteCommerceBuy(Drag, Val(cantidad.Text))
        Else
            AddtoRichTextBox frmMain.RecTxt, "No tenés suficiente oro.", 2, 51, 223, 1, 1

        End If

    End If

    
    Exit Sub

InvComNpc_ItemDropped_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.InvComNpc_ItemDropped", Erl)
    Resume Next
    
End Sub

Private Sub tmrNumber_Timer()
    
    On Error GoTo tmrNumber_Timer_Err
    

    Const MIN_NUMBER = 1

    Const MAX_NUMBER = 10000

    cantidad = cantidad + m_Increment

    If cantidad < MIN_NUMBER Then
        cantidad = MIN_NUMBER
    ElseIf cantidad > MAX_NUMBER Then
        cantidad = MAX_NUMBER

    End If

    cantidad.Text = format$(cantidad)
    
    If m_Interval > 1 Then
        m_Interval = m_Interval - 1
        tmrNumber.Interval = m_Interval

    End If

    
    Exit Sub

tmrNumber_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.tmrNumber_Timer", Erl)
    Resume Next
    
End Sub
