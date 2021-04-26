VERSION 5.00
Begin VB.Form frmBancoObj 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Cadena de finanzas Goliath"
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   481
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   544
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
         Size            =   8.25
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
      ScaleHeight     =   244
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   459
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
   Begin VB.Image cmdMasMenos 
      Height          =   315
      Index           =   0
      Left            =   3195
      Tag             =   "0"
      Top             =   6525
      Width           =   315
   End
   Begin VB.Image cmdMasMenos 
      Height          =   315
      Index           =   1
      Left            =   4650
      Tag             =   "0"
      Top             =   6510
      Width           =   315
   End
   Begin VB.Label lblcosto 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
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
         Size            =   6.75
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
   Begin VB.Image Image1 
      Height          =   420
      Index           =   1
      Left            =   5505
      Tag             =   "0"
      Top             =   6465
      Width           =   1830
   End
   Begin VB.Image Image1 
      Height          =   420
      Index           =   0
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

' Declaro los inventarios acá para manejar el evento drop
Public WithEvents InvBankUsu As clsGrapchicalInventory ' Inventario del usuario visible en la bóveda
Attribute InvBankUsu.VB_VarHelpID = -1

Public WithEvents InvBoveda  As clsGrapchicalInventory ' Inventario de la bóveda
Attribute InvBoveda.VB_VarHelpID = -1

Private Sub moverForm()
    
    On Error GoTo moverForm_Err
    

    Dim res As Long

    ReleaseCapture
    res = SendMessage(Me.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)

    
    Exit Sub

moverForm_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoObj.moverForm", Erl)
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
    Call RegistrarError(Err.number, Err.Description, "frmBancoObj.cantidad_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub cmdMasMenos_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdMasMenos_MouseDown_Err
    

    Call Sound.Sound_Play(SND_CLICK)

    Select Case Index

        Case 0
            cmdMasMenos(Index).Picture = LoadInterface("boton-sm-menos-off.bmp")
            cmdMasMenos(Index).Tag = "1"
            cantidad.Text = str((Val(cantidad.Text) - 1))
            m_Increment = -1

        Case 1
            cmdMasMenos(Index).Picture = LoadInterface("boton-sm-mas-off.bmp")
            cmdMasMenos(Index).Tag = "1"
            m_Increment = 1

    End Select

    tmrNumber.Interval = 30
    tmrNumber.Enabled = True

    
    Exit Sub

cmdMasMenos_MouseDown_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoObj.cmdMasMenos_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub cmdMasMenos_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdMasMenos_MouseMove_Err
    

    Select Case Index

        Case 0

            If cmdMasMenos(Index).Tag = "0" Then
                cmdMasMenos(Index).Picture = LoadInterface("boton-sm-menos-over.bmp")
                cmdMasMenos(Index).Tag = "1"

            End If

        Case 1

            If cmdMasMenos(Index).Tag = "0" Then
                cmdMasMenos(Index).Picture = LoadInterface("boton-sm-mas-over.bmp")
                cmdMasMenos(Index).Tag = "1"

            End If

    End Select

    
    Exit Sub

cmdMasMenos_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoObj.cmdMasMenos_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub cmdMasMenos_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdMasMenos_MouseUp_Err
    
    Call Form_MouseMove(Button, Shift, x, y)
    tmrNumber.Enabled = False

    
    Exit Sub

cmdMasMenos_MouseUp_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoObj.cmdMasMenos_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoObj.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    cantidad.BackColor = RGB(18, 19, 13)

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoObj.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Image1_Click(Index As Integer)
    
    On Error GoTo Image1_Click_Err
    
    Call Sound.Sound_Play(SND_CLICK)

    If Not IsNumeric(cantidad.Text) Then Exit Sub
    If Val(cantidad.Text) <= 0 Then Exit Sub

    Select Case Index

        Case 0
            'frmBancoObj.List1(0).SetFocus
            'LastIndex1 = List1(0).ListIndex
            LasActionBuy = True
            'Call WriteBankExtractItem(InvBoveda.SelectedItem, cantidad.Text, 1)
        
            If InvBoveda.SelectedItem <= 0 Then Exit Sub
            Call WriteBankExtractItem(InvBoveda.SelectedItem, min(Val(cantidad.Text), InvBoveda.Amount(InvBoveda.SelectedItem)), 0)

        Case 1
            LasActionBuy = False

            If InvBankUsu.SelectedItem <= 0 Then Exit Sub
            Call WriteBankDeposit(InvBankUsu.SelectedItem, min(Val(cantidad.Text), InvBankUsu.Amount(InvBankUsu.SelectedItem)), 0)

    End Select

    
    Exit Sub

Image1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoObj.Image1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    moverForm

    If Image1(0).Tag = "1" Then
        Image1(0).Picture = Nothing
        Image1(0).Tag = "0"

    End If

    If Image1(1).Tag = "1" Then
        Image1(1).Picture = Nothing
        Image1(1).Tag = "0"

    End If

    If cmdMasMenos(0).Tag = "1" Then
        cmdMasMenos(0).Picture = Nothing
        cmdMasMenos(0).Tag = "0"

    End If

    If cmdMasMenos(1).Tag = "1" Then
        cmdMasMenos(1).Picture = Nothing
        cmdMasMenos(1).Tag = "0"

    End If

    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoObj.Form_MouseMove", Erl)
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
    Call RegistrarError(Err.number, Err.Description, "frmBancoObj.cantidad_Change", Erl)
    Resume Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error GoTo Form_Unload_Err
    
    Call Sound.Sound_Play(SND_CLICK)
    Call WriteBankEnd

    
    Exit Sub

Form_Unload_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoObj.Form_Unload", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image1_MouseDown_Err
    

    If Index = 0 Then
        Image1(0).Picture = LoadInterface("boton-retirar-ES-off.bmp")
        Image1(0).Tag = "0"
    Else
        Image1(1).Picture = LoadInterface("boton-depositar-ES-off.bmp")
        Image1(1).Tag = "0"

    End If

    
    Exit Sub

Image1_MouseDown_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoObj.Image1_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image1_MouseMove_Err
    

    If Index = 0 Then
        If Image1(0).Tag = "0" Then
            Image1(0).Picture = LoadInterface("boton-retirar-ES-over.bmp")
            Image1(0).Tag = "1"

        End If

    Else
    
        If Image1(1).Tag = "0" Then
            Image1(1).Picture = LoadInterface("boton-depositar-ES-default.bmp")
            Image1(1).Tag = "1"

        End If

    End If

    
    Exit Sub

Image1_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoObj.Image1_MouseMove", Erl)
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
    Call RegistrarError(Err.number, Err.Description, "frmBancoObj.interface_Click", Erl)
    Resume Next
    
End Sub

Private Sub interface_DblClick()
    
    On Error GoTo interface_DblClick_Err
    

    ' Nos aseguramos que lo último que cliqueó fue el inventario
    If Not InvBankUsu.ClickedInside Then Exit Sub
    
    If Not InvBankUsu.IsItemSelected Then Exit Sub

    ' Hacemos acción del doble clic correspondiente
    Dim ObjType As Byte

    ObjType = ObjData(InvBankUsu.OBJIndex(InvBankUsu.SelectedItem)).ObjType
    
    If UserMeditar Then Exit Sub
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
    Select Case ObjType

        Case eObjType.otArmadura, eObjType.otESCUDO, eObjType.otmagicos, eObjType.otFlechas, eObjType.otCASCO, eObjType.otNudillos, eObjType.otAnillos
            Call WriteEquipItem(InvBankUsu.SelectedItem)
            
        Case eObjType.otWeapon

            If ObjData(InvBankUsu.OBJIndex(InvBankUsu.SelectedItem)).proyectil = 1 And InvBankUsu.Equipped(InvBankUsu.SelectedItem) Then
                Call WriteUseItem(InvBankUsu.SelectedItem)
            Else
                Call WriteEquipItem(InvBankUsu.SelectedItem)

            End If
            
        Case eObjType.OtHerramientas

            If InvBankUsu.Equipped(InvBankUsu.SelectedItem) Then
                Call WriteUseItem(InvBankUsu.SelectedItem)
            Else
                Call WriteEquipItem(InvBankUsu.SelectedItem)

            End If
             
        Case Else
            Call WriteUseItem(InvBankUsu.SelectedItem)

    End Select

    
    Exit Sub

interface_DblClick_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoObj.interface_DblClick", Erl)
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
    Call RegistrarError(Err.number, Err.Description, "frmBancoObj.interface_KeyDown", Erl)
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
    Call RegistrarError(Err.number, Err.Description, "frmBancoObj.InvBoveda_ItemDropped", Erl)
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
    Call RegistrarError(Err.number, Err.Description, "frmBancoObj.InvBankUsu_ItemDropped", Erl)
    Resume Next
    
End Sub

Private Sub salir_Click()
    
    On Error GoTo salir_Click_Err
    
    Unload Me

    
    Exit Sub

salir_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmBancoObj.salir_Click", Erl)
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
    Call RegistrarError(Err.number, Err.Description, "frmBancoObj.tmrNumber_Timer", Erl)
    Resume Next
    
End Sub
