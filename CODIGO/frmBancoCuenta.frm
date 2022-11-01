VERSION 5.00
Begin VB.Form frmBancoCuenta 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7152
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8148
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   596
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   679
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   620
      MousePointer    =   99  'Custom
      ScaleHeight     =   305
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   574
      TabIndex        =   1
      TabStop         =   0   'False
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
      Left            =   3690
      TabIndex        =   0
      Text            =   "1"
      Top             =   6555
      Width           =   810
   End
   Begin VB.Timer tmrNumber 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   420
      Index           =   0
      Left            =   825
      Tag             =   "0"
      Top             =   6465
      Width           =   1830
   End
   Begin VB.Image Image1 
      Height          =   420
      Index           =   1
      Left            =   5505
      Tag             =   "0"
      Top             =   6465
      Width           =   1830
   End
   Begin VB.Label lbldesc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   5910
      Width           =   3135
   End
   Begin VB.Label lblnombre 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Vacio)"
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
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   5700
      Width           =   3135
   End
   Begin VB.Label lblcosto 
      AutoSize        =   -1  'True
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
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Image cmdMasMenos 
      Height          =   315
      Index           =   1
      Left            =   4650
      Tag             =   "0"
      Top             =   6510
      Width           =   315
   End
   Begin VB.Image cmdMasMenos 
      Height          =   315
      Index           =   0
      Left            =   3240
      Tag             =   "0"
      Top             =   6525
      Width           =   315
   End
   Begin VB.Image salir 
      Height          =   375
      Left            =   7680
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmBancoCuenta"
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

Public LastIndex1            As Integer

Public LasActionBuy          As Boolean

Private m_Number             As Integer

Private m_Increment          As Integer

Private m_Interval           As Integer

Private cBotonRetirar As clsGraphicalButton
Private cBotonDepositar As clsGraphicalButton
Private cBotonMas As clsGraphicalButton
Private cBotonMenos As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton

' Declaro los inventarios acá para manejar el evento drop
Public WithEvents InvBankUsuCuenta As clsGrapchicalInventory ' Inventario del usuario visible en la bóveda
Attribute InvBankUsuCuenta.VB_VarHelpID = -1

Public WithEvents InvBovedaCuenta  As clsGrapchicalInventory ' Inventario de la bóveda
Attribute InvBovedaCuenta.VB_VarHelpID = -1

Private Sub MoverForm()
    
    On Error GoTo moverForm_Err
    

    Dim res As Long

    ReleaseCapture
    res = SendMessage(Me.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)

    
    Exit Sub

moverForm_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoCuenta.moverForm", Erl)
    Resume Next
    
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
    Call RegistrarError(Err.number, Err.Description, "frmBancoCuenta.cantidad_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub cmdMasMenos_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdMasMenos_MouseDown_Err
    

    Call Sound.Sound_Play(SND_CLICK)

    Select Case Index

        Case 0
            cmdMasMenos(Index).Picture = LoadInterface("boton-sm-menos-off.bmp")
            cantidad.Text = str((Val(cantidad.Text) - 1))
            m_Increment = -1

        Case 1
            cmdMasMenos(Index).Picture = LoadInterface("boton-sm-mas-off.bmp")
            m_Increment = 1

    End Select

    tmrNumber.Interval = 30
    tmrNumber.Enabled = True

    
    Exit Sub

cmdMasMenos_MouseDown_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoCuenta.cmdMasMenos_MouseDown", Erl)
    Resume Next
    
End Sub


Private Sub cmdMasMenos_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdMasMenos_MouseUp_Err
    
    'Call Form_MouseMove(Button, Shift, x, y)
    tmrNumber.Enabled = False

    
    Exit Sub

cmdMasMenos_MouseUp_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoCuenta.cmdMasMenos_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoCuenta.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Me.Picture = LoadInterface("banco.bmp")
    
    Call FormParser.Parse_Form(Me)
    cantidad.BackColor = RGB(18, 19, 13)
    If Me.Visible Then
        Call LoadButtons
    End If
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoCuenta.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub LoadButtons()
        
    Set cBotonRetirar = New clsGraphicalButton
    Set cBotonDepositar = New clsGraphicalButton
    Set cBotonMas = New clsGraphicalButton
    Set cBotonMenos = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton

    Call cBotonCerrar.Initialize(salir, "boton-cerrar-default.bmp", _
                                                "boton-cerrar-over.bmp", _
                                                "boton-cerrar-off.bmp", Me)
                                                                                            
    Call cBotonMas.Initialize(cmdMasMenos(0), "boton-sm-mas-default.bmp", _
                                                "boton-sm-mas-over.bmp", _
                                                "boton-sm-mas-off.bmp", Me)
                                                
    Call cBotonMenos.Initialize(cmdMasMenos(1), "boton-sm-menos-default.bmp", _
                                                "boton-sm-menos-over.bmp", _
                                                "boton-sm-menos-off.bmp", Me)
                                                
    Call cBotonRetirar.Initialize(Image1(0), "boton-retirar-default.bmp", _
                                                "boton-retirar-over.bmp", _
                                                "boton-retirar-off.bmp", Me)
    
    Call cBotonDepositar.Initialize(Image1(1), "boton-depositar-default.bmp", _
                                                "boton-depositar-over.bmp", _
                                                "boton-depositar-off.bmp", Me)
                                                
    
End Sub
Private Sub Image1_Click(Index As Integer)
    
    On Error GoTo Image1_Click_Err
    
    Call Sound.Sound_Play(SND_CLICK)

    If Not IsNumeric(cantidad.Text) Then Exit Sub
    If Val(cantidad.Text) <= 0 Then Exit Sub

    Select Case Index

        Case 0
            LasActionBuy = True
            
            If InvBovedaCuenta.SelectedItem <= 0 Then Exit Sub
            Call WriteCuentaExtractItem(InvBovedaCuenta.SelectedItem, min(Val(cantidad.Text), InvBovedaCuenta.Amount(InvBovedaCuenta.SelectedItem)), 0)

        Case 1
            LasActionBuy = False

            If InvBankUsuCuenta.SelectedItem <= 0 Then Exit Sub
            Call WriteCuentaDeposit(InvBankUsuCuenta.SelectedItem, min(Val(cantidad.Text), InvBankUsuCuenta.Amount(InvBankUsuCuenta.SelectedItem)), 0)
    End Select

    
    Exit Sub

Image1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoCuenta.Image1_Click", Erl)
    Resume Next
    
End Sub

Private Sub cantidad_Change()
    
    On Error GoTo cantidad_Change_Err
    

    If Val(cantidad.Text) < 1 Then
        cantidad.Text = 1

    End If
    
    If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
        cantidad.Text = MAX_INVENTORY_OBJS
        cantidad.SelStart = Len(cantidad.Text)

    End If

    
    Exit Sub

cantidad_Change_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoCuenta.cantidad_Change", Erl)
    Resume Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error GoTo Form_Unload_Err
    
    Call Sound.Sound_Play(SND_CLICK)
    If Not Protocol_Writes.writer_is_nothing Then
        Call WriteBankEnd
    End If
    
    Exit Sub

Form_Unload_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoCuenta.Form_Unload", Erl)
    Resume Next
    
End Sub


Private Sub interface_Click()
    
    On Error GoTo interface_Click_Err
    
    
    If InvBovedaCuenta.ClickedInside Then
        ' Cliqueé en la bóveda, deselecciono el inventario
        Call InvBankUsuCuenta.SeleccionarItem(0)
        
    ElseIf InvBankUsuCuenta.ClickedInside Then
        ' Cliqueé en el inventario, deselecciono la bóveda
        Call InvBovedaCuenta.SeleccionarItem(0)

    End If

    
    Exit Sub

interface_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoCuenta.interface_Click", Erl)
    Resume Next
    
End Sub

Private Sub interface_DblClick()
    
    On Error GoTo interface_DblClick_Err
    

    ' Nos aseguramos que lo último que cliqueó fue el inventario
    If Not InvBankUsuCuenta.ClickedInside Then Exit Sub
    
    If Not InvBankUsuCuenta.IsItemSelected Then Exit Sub

    ' Hacemos acción del doble clic correspondiente
    Dim ObjType As Byte

    ObjType = ObjData(InvBankUsuCuenta.OBJIndex(InvBankUsuCuenta.SelectedItem)).ObjType
    
    If UserMeditar Then Exit Sub
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
    Select Case ObjType

        Case eObjType.otArmadura, eObjType.otESCUDO, eObjType.otmagicos, eObjType.otFlechas, eObjType.otCASCO, eObjType.otNudillos, eObjType.otAnillos
            Call WriteEquipItem(InvBankUsuCuenta.SelectedItem)
            
        Case eObjType.otWeapon

            If ObjData(InvBankUsuCuenta.OBJIndex(InvBankUsuCuenta.SelectedItem)).proyectil = 1 And InvBankUsuCuenta.Equipped(InvBankUsuCuenta.SelectedItem) Then
                Call WriteUseItem(InvBankUsuCuenta.SelectedItem)
            Else
                Call WriteEquipItem(InvBankUsuCuenta.SelectedItem)

            End If
            
        Case eObjType.OtHerramientas

            If InvBankUsuCuenta.Equipped(InvBankUsuCuenta.SelectedItem) Then
                Call WriteUseItem(InvBankUsuCuenta.SelectedItem)
            Else
                Call WriteEquipItem(InvBankUsuCuenta.SelectedItem)

            End If
             
        Case Else
            Call WriteUseItem(InvBankUsuCuenta.SelectedItem)

    End Select

    
    Exit Sub

interface_DblClick_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoCuenta.interface_DblClick", Erl)
    Resume Next
    
End Sub

Private Sub interface_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo interface_KeyDown_Err
    

    ' Referencia temporal al inventario que corresponda
    Dim CurrentInventory As clsGrapchicalInventory

    If InvBovedaCuenta.ClickedInside Then
        Set CurrentInventory = InvBovedaCuenta
    ElseIf InvBankUsuCuenta.ClickedInside Then
        Set CurrentInventory = InvBankUsuCuenta
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
    Call RegistrarError(Err.number, Err.Description, "frmBancoCuenta.interface_KeyDown", Erl)
    Resume Next
    
End Sub

Private Sub InvBovedaCuenta_ItemDropped(ByVal Drag As Integer, ByVal Drop As Integer, ByVal x As Integer, ByVal y As Integer)
    
    On Error GoTo InvBovedaCuenta_ItemDropped_Err
    

    ' Si lo soltó dentro de la bóveda
    If Drop > 0 Then
        ' Movemos el item dentro de la bóveda
        'Call WriteBovedaItemMove(Drag, Drop)
    Else
        Drop = InvBankUsuCuenta.GetSlot(x, y)

        ' Si lo soltó dentro del inventario
        If Drop > 0 Then
            ' Retiramos el item
            Call WriteCuentaExtractItem(Drag, min(Val(cantidad.Text), InvBovedaCuenta.Amount(InvBovedaCuenta.SelectedItem)), Drop)

        End If

    End If

    
    Exit Sub

InvBovedaCuenta_ItemDropped_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoCuenta.InvBovedaCuenta_ItemDropped", Erl)
    Resume Next
    
End Sub

Private Sub InvBankUsuCuenta_ItemDropped(ByVal Drag As Integer, ByVal Drop As Integer, ByVal x As Integer, ByVal y As Integer)
    
    On Error GoTo InvBankUsuCuenta_ItemDropped_Err
    

    ' Si lo soltó dentro del mismo inventario
    If Drop > 0 Then
        ' Si el slot de drag es distinto del de drop
        If Drag <> Drop Then
            ' Movemos el item dentro del inventario
            Call WriteItemMove(Drag, Drop)
        End If
    Else
        Drop = InvBovedaCuenta.GetSlot(x, y)

        ' Si lo soltó dentro de la bóveda
        If Drop > 0 Then
            ' Depositamos el item
            Call WriteCuentaDeposit(Drag, min(Val(cantidad.Text), InvBankUsuCuenta.Amount(InvBankUsuCuenta.SelectedItem)), Drop)

        End If

    End If

    
    Exit Sub

InvBankUsuCuenta_ItemDropped_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoCuenta.InvBankUsuCuenta_ItemDropped", Erl)
    Resume Next
    
End Sub

Private Sub salir_Click()
    
    On Error GoTo salir_Click_Err
    
    Unload Me

    
    Exit Sub

salir_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoCuenta.salir_Click", Erl)
    Resume Next
    
End Sub

Private Sub tmrNumber_Timer()
    
    On Error GoTo tmrNumber_Timer_Err
    

    Const MIN_NUMBER = 1

    Const MAX_NUMBER = 10000

    m_Number = m_Number + m_Increment

    If m_Number < MIN_NUMBER Then
        m_Number = MIN_NUMBER
    ElseIf m_Number > MAX_NUMBER Then
        m_Number = MAX_NUMBER

    End If

    cantidad.Text = format$(m_Number)
    
    If m_Interval > 1 Then
        m_Interval = m_Interval - 1
        tmrNumber.Interval = m_Interval

    End If

    
    Exit Sub

tmrNumber_Timer_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoCuenta.tmrNumber_Timer", Erl)
    Resume Next
    
End Sub


