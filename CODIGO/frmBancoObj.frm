VERSION 5.00
Begin VB.Form frmBancoObj 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Cadena de finanzas Goliath"
   ClientHeight    =   7212
   ClientLeft      =   0
   ClientTop       =   -72
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   601
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrNumber 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   0
      Top             =   0
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
      TabIndex        =   4
      Text            =   "1"
      Top             =   6555
      Width           =   810
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
      Left            =   630
      MousePointer    =   99  'Custom
      ScaleHeight     =   305
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   574
      TabIndex        =   0
      Top             =   1605
      Width           =   6885
   End
   Begin VB.Image salir 
      Height          =   375
      Left            =   7680
      Top             =   0
      Width           =   495
   End
   Begin VB.Image cmdMenos 
      Height          =   315
      Left            =   3195
      Tag             =   "0"
      Top             =   6525
      Width           =   315
   End
   Begin VB.Image cmdMas 
      Height          =   315
      Left            =   4650
      Tag             =   "0"
      Top             =   6480
      Width           =   315
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
      TabIndex        =   3
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label lblnombre 
      Alignment       =   2  'Center
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
      TabIndex        =   2
      Top             =   5700
      Width           =   3135
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
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   5910
      Width           =   3135
   End
   Begin VB.Image cmdDepositar 
      Height          =   420
      Left            =   5505
      Tag             =   "0"
      Top             =   6465
      Width           =   1830
   End
   Begin VB.Image cmdRetirar 
      Height          =   420
      Left            =   825
      Tag             =   "0"
      Top             =   6465
      Width           =   1830
   End
End
Attribute VB_Name = "frmBancoObj"
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

Private cBotonRetirar As clsGraphicalButton
Private cBotonDepositar As clsGraphicalButton
Private cBotonMas As clsGraphicalButton
Private cBotonMenos As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton

Const WM_SYSCOMMAND As Long = &H112&

Const MOUSE_MOVE    As Long = &HF012&

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Public LastIndex1            As Integer

Public LasActionBuy          As Boolean

Private m_Number             As Integer

Private m_Increment          As Integer

Private m_Interval           As Integer


' Declaro los inventarios acá para manejar el evento drop
Public WithEvents InvBankUsu As clsGrapchicalInventory ' Inventario del usuario visible en la bóveda
Attribute InvBankUsu.VB_VarHelpID = -1

Public WithEvents InvBoveda  As clsGrapchicalInventory ' Inventario de la bóveda
Attribute InvBoveda.VB_VarHelpID = -1

Private Sub MoverForm()
    
    On Error GoTo moverForm_Err
    

    Dim res As Long

    ReleaseCapture
    res = SendMessage(Me.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)

    
    Exit Sub

moverForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmBancoObj.moverForm", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "frmBancoObj.cantidad_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub cmdMas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_Increment = 1
    tmrNumber.Interval = 30
    tmrNumber.Enabled = True
    Exit Sub
End Sub

Private Sub cmdMas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tmrNumber.Enabled = False
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmBancoObj.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err

    Me.Picture = LoadInterface("banco.bmp")
    
    Call FormParser.Parse_Form(Me)
    cantidad.BackColor = RGB(18, 19, 13)

    Call LoadButtons

    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmBancoObj.Form_Load", Erl)
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
                                                                                            
    Call cBotonMas.Initialize(cmdMas, "boton-sm-mas-default.bmp", _
                                                "boton-sm-mas-over.bmp", _
                                                "boton-sm-mas-off.bmp", Me)
                                                
    Call cBotonMenos.Initialize(cmdMenos, "boton-sm-menos-default.bmp", _
                                                "boton-sm-menos-over.bmp", _
                                                "boton-sm-menos-off.bmp", Me)
                                                
    Call cBotonRetirar.Initialize(cmdRetirar, "boton-retirar-default.bmp", _
                                                "boton-retirar-over.bmp", _
                                                "boton-retirar-off.bmp", Me)
    
    Call cBotonDepositar.Initialize(cmdDepositar, "boton-depositar-default.bmp", _
                                                "boton-depositar-over.bmp", _
                                                "boton-depositar-off.bmp", Me)
                                                
    
End Sub
Private Sub cmdDepositar_Click()
    If Not IsNumeric(cantidad.Text) Then Exit Sub
    If Val(cantidad.Text) <= 0 Then Exit Sub
    
    LasActionBuy = False

    If InvBankUsu.SelectedItem <= 0 Then Exit Sub
    Call WriteBankDeposit(InvBankUsu.SelectedItem, min(Val(cantidad.Text), InvBankUsu.Amount(InvBankUsu.SelectedItem)), 0)
End Sub

Private Sub cmdRetirar_Click()
    If Not IsNumeric(cantidad.Text) Then Exit Sub
    If Val(cantidad.Text) <= 0 Then Exit Sub

    LasActionBuy = True

    If InvBoveda.SelectedItem <= 0 Then Exit Sub
    Call WriteBankExtractItem(InvBoveda.SelectedItem, min(Val(cantidad.Text), InvBoveda.Amount(InvBoveda.SelectedItem)), 0)
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
    Call RegistrarError(Err.Number, Err.Description, "frmBancoObj.cantidad_Change", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "frmBancoObj.Form_Unload", Erl)
    Resume Next
    
End Sub

Private Sub interface_Click()
    
    On Error GoTo interface_Click_Err
    
    
    If InvBoveda.ClickedInside Then
        ' Cliqueé en la bóveda, deselecciono el inventario
        Call InvBankUsu.SeleccionarItem(0)
        
    ElseIf InvBankUsu.ClickedInside Then
        ' Cliqueé en el inventario, deselecciono la bóveda
        Call InvBoveda.SeleccionarItem(0)

    End If

    
    Exit Sub

interface_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmBancoObj.interface_Click", Erl)
    Resume Next
    
End Sub

Private Sub interface_DblClick()

    On Error GoTo interface_DblClick_Err

    If InvBankUsu.ClickedInside And InvBankUsu.IsItemSelected Then
        Call WriteBankDeposit(InvBankUsu.SelectedItem, InvBankUsu.Amount(InvBankUsu.SelectedItem), 0)
    ElseIf InvBoveda.ClickedInside And InvBoveda.IsItemSelected Then
        Call WriteBankExtractItem(InvBoveda.SelectedItem, InvBoveda.Amount(InvBoveda.SelectedItem), 0)
    End If

    
    Exit Sub

interface_DblClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmBancoObj.interface_DblClick", Erl)
    Resume Next
    
End Sub

Private Sub interface_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo interface_KeyDown_Err
    

    ' Referencia temporal al inventario que corresponda
    Dim CurrentInventory As clsGrapchicalInventory

    If InvBoveda.ClickedInside Then
        Set CurrentInventory = InvBoveda
    ElseIf InvBankUsu.ClickedInside Then
        Set CurrentInventory = InvBankUsu
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
    Call RegistrarError(Err.Number, Err.Description, "frmBancoObj.interface_KeyDown", Erl)
    Resume Next
    
End Sub

Private Sub InvBoveda_ItemDropped(ByVal Drag As Integer, ByVal Drop As Integer, ByVal x As Integer, ByVal y As Integer)
    
    On Error GoTo InvBoveda_ItemDropped_Err
    

    ' Si lo soltó dentro de la bóveda
    If Drop > 0 Then
        ' Movemos el item dentro de la bóveda
        Call WriteBovedaItemMove(Drag, Drop)
    Else
        Drop = InvBankUsu.GetSlot(x, y)

        ' Si lo soltó dentro del inventario
        If Drop > 0 Then
            ' Retiramos el item
            Call WriteBankExtractItem(Drag, min(Val(cantidad.Text), InvBoveda.Amount(InvBoveda.SelectedItem)), Drop)

        End If

    End If

    
    Exit Sub

InvBoveda_ItemDropped_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmBancoObj.InvBoveda_ItemDropped", Erl)
    Resume Next
    
End Sub

Private Sub InvBankUsu_ItemDropped(ByVal Drag As Integer, ByVal Drop As Integer, ByVal x As Integer, ByVal y As Integer)
    
    On Error GoTo InvBankUsu_ItemDropped_Err
    

    ' Si lo soltó dentro del mismo inventario
    If Drop > 0 Then
        If Drag <> Drop Then
            ' Movemos el item dentro del inventario
            Call WriteItemMove(Drag, Drop)
        End If
    Else
        Drop = InvBoveda.GetSlot(x, y)

        ' Si lo soltó dentro de la bóveda
        If Drop > 0 Then
            ' Depositamos el item
            Call WriteBankDeposit(Drag, min(Val(cantidad.Text), InvBankUsu.Amount(InvBankUsu.SelectedItem)), Drop)

        End If

    End If

    
    Exit Sub

InvBankUsu_ItemDropped_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmBancoObj.InvBankUsu_ItemDropped", Erl)
    Resume Next
    
End Sub

Private Sub salir_Click()
    
    On Error GoTo salir_Click_Err
    
    Unload Me

    
    Exit Sub

salir_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmBancoObj.salir_Click", Erl)
    Resume Next
    
End Sub

Private Sub tmrNumber_Timer()
    
    On Error GoTo tmrNumber_Timer_Err
    

    Const MIN_NUMBER = 1

    Const MAX_NUMBER = 10000

    cantidad.Text = Val(cantidad.Text) + m_Increment

    If Val(cantidad.Text) < MIN_NUMBER Then
        cantidad.Text = MIN_NUMBER
    ElseIf Val(cantidad.Text) > MAX_NUMBER Then
        cantidad.Text = MAX_NUMBER

    End If

    cantidad.Text = format$(Val(cantidad.Text))
    
    If m_Interval > 1 Then
        m_Interval = m_Interval - 1
        tmrNumber.Interval = m_Interval

    End If

    
    Exit Sub

tmrNumber_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmBancoObj.tmrNumber_Timer", Erl)
    Resume Next
    
End Sub
