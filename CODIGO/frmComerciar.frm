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
   ScaleHeight     =   476
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   544
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
   Begin VB.Image salir 
      Height          =   375
      Left            =   7680
      Top             =   0
      Width           =   495
   End
   Begin VB.Image cmdMasMenos 
      Height          =   315
      Index           =   1
      Left            =   4650
      Tag             =   "1"
      Top             =   6510
      Width           =   315
   End
   Begin VB.Image cmdMasMenos 
      Height          =   315
      Index           =   0
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
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   5820
      Width           =   3135
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
      TabIndex        =   3
      Top             =   5610
      Width           =   3135
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
      TabIndex        =   2
      Top             =   5550
      Width           =   1215
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
Attribute VB_Name = "frmComerciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const WM_SYSCOMMAND As Long = &H112&

Const MOUSE_MOVE    As Long = &HF012&

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

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

Private Sub MoverForm()
    
    On Error GoTo moverForm_Err
    

    Dim res As Long

    ReleaseCapture
    res = SendMessage(Me.hWnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)

    
    Exit Sub

moverForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.moverForm", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.cantidad_KeyPress", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.cmdMasMenos_MouseDown", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.cmdMasMenos_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub cmdMasMenos_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdMasMenos_MouseUp_Err
    
    Call Form_MouseMove(Button, Shift, x, y)
    tmrNumber.Enabled = False

    
    Exit Sub

cmdMasMenos_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.cmdMasMenos_MouseUp", Erl)
    Resume Next
    
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

Private Sub Image1_Click(Index As Integer)
    
    On Error GoTo Image1_Click_Err
    
    Call Sound.Sound_Play(SND_CLICK)
    
    If Not IsNumeric(cantidad.Text) Then Exit Sub
    If Val(cantidad.Text) <= 0 Then Exit Sub

    Select Case Index

        Case 0

            If InvComNpc.SelectedItem <= 0 Then Exit Sub
 
            LasActionBuy = True

            If UserGLD >= InvComNpc.Valor(InvComNpc.SelectedItem) * Val(cantidad) Then
                Call WriteCommerceBuy(InvComNpc.SelectedItem, cantidad.Text)
            Else
                AddtoRichTextBox frmMain.RecTxt, "No tenés suficiente oro.", 2, 51, 223, 1, 1

            End If
       
        Case 1

            If InvComUsu.SelectedItem <= 0 Then Exit Sub
            
            LasActionBuy = False
            
            Call WriteCommerceSell(InvComUsu.SelectedItem, min(Val(cantidad.Text), InvComUsu.Amount(InvComUsu.SelectedItem)))

    End Select
    
    
    Exit Sub

Image1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.Image1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    cantidad.BackColor = RGB(18, 19, 13)

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    MoverForm

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
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.Form_MouseMove", Erl)
    Resume Next
    
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
    
    Call WriteCommerceEnd

    
    Exit Sub

Form_Unload_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.Form_Unload", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image1_MouseDown_Err
    

    If Index = 0 Then
        Image1(0).Picture = LoadInterface("boton-comprar-ES-off.bmp")
        Image1(0).Tag = "0"
    Else
        Image1(1).Picture = LoadInterface("boton-vender-ES-off.bmp")
        Image1(1).Tag = "0"

    End If

    
    Exit Sub

Image1_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.Image1_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image1_MouseMove_Err
    

    If Index = 0 Then
        If Image1(0).Tag = "0" Then
            Image1(0).Picture = LoadInterface("boton-comprar-ES-over.bmp")
            Image1(0).Tag = "1"

        End If

    Else
        
        If Image1(1).Tag = "0" Then
            Image1(1).Picture = LoadInterface("boton-vender-over.bmp")
            Image1(1).Tag = "1"

        End If

    End If

    
    Exit Sub

Image1_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.Image1_MouseMove", Erl)
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

        ObjType = ObjData(InvComUsu.objIndex(InvComUsu.SelectedItem)).ObjType
        
        If UserMeditar Then Exit Sub
        If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
        
        Select Case ObjType

            Case eObjType.otArmadura, eObjType.otESCUDO, eObjType.otmagicos, eObjType.otFlechas, eObjType.otCASCO, eObjType.otNudillos, eObjType.otAnillos
                Call WriteEquipItem(InvComUsu.SelectedItem)
                
            Case eObjType.otWeapon

                If ObjData(InvComUsu.objIndex(InvComUsu.SelectedItem)).proyectil = 1 And InvComUsu.Equipped(InvComUsu.SelectedItem) Then
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

Private Sub salir_Click()
    
    On Error GoTo salir_Click_Err
    
    Unload Me

    
    Exit Sub

salir_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.salir_Click", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "frmComerciar.tmrNumber_Timer", Erl)
    Resume Next
    
End Sub
