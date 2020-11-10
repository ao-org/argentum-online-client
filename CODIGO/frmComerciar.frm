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

Private Sub moverForm()

    Dim res As Long

    ReleaseCapture
    res = SendMessage(Me.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)

End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)

    If (KeyAscii = 27) Then
        Unload Me

    End If

    If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0

        End If

    End If

End Sub

Private Sub cmdMasMenos_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

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

End Sub

Private Sub cmdMasMenos_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

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

End Sub

Private Sub cmdMasMenos_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Form_MouseMove(Button, Shift, x, y)
    tmrNumber.Enabled = False

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If (KeyAscii = 27) Then
        Unload Me

    End If

End Sub

Private Sub Image1_Click(Index As Integer)
    Call Sound.Sound_Play(SND_CLICK)
    
    If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub

    Select Case Index

        Case 0

            If InvComNpc.SelectedItem <= 0 Then Exit Sub
 
            LasActionBuy = True

            If UserGLD >= InvComNpc.Valor(InvComNpc.SelectedItem) * Val(cantidad) Then
                Call WriteCommerceBuy(InvComNpc.SelectedItem, cantidad.Text)
            Else
                AddtoRichTextBox frmmain.RecTxt, "No tenés suficiente oro.", 2, 51, 223, 1, 1

            End If
       
        Case 1

            If InvComUsu.SelectedItem <= 0 Then Exit Sub
            
            LasActionBuy = False
            
            Call WriteCommerceSell(InvComUsu.SelectedItem, min(Val(cantidad.Text), InvComUsu.Amount(InvComUsu.SelectedItem)))

    End Select
    
End Sub

Private Sub Form_Load()
    Call FormParser.Parse_Form(Me)
    cantidad.BackColor = RGB(18, 19, 13)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
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

End Sub

Private Sub addRemove_Click(Index As Integer)
    Call Sound.Sound_Play(SND_CLICK)

    Select Case Index

        Case 0
            cantidad = cantidad - 1

        Case 1
            cantidad = cantidad + 1

    End Select

End Sub

Private Sub cantidad_Change()

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

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call WriteCommerceEnd

End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If Index = 0 Then
        Image1(0).Picture = LoadInterface("boton-comprar-ES-off.bmp")
        Image1(0).Tag = "0"
    Else
        Image1(1).Picture = LoadInterface("boton-vender-ES-off.bmp")
        Image1(1).Tag = "0"

    End If

End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

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

End Sub

Private Sub interface_Click()
    
    If InvComNpc.ClickedInside Then
        ' Cliqueé en la tienda, deselecciono el inventario
        Call InvComUsu.SeleccionarItem(0)
        
    ElseIf InvComUsu.ClickedInside Then
        ' Cliqueé en el inventario, deselecciono la tienda
        Call InvComNpc.SeleccionarItem(0)

    End If

End Sub

Private Sub interface_DblClick()

    If InvComNpc.ClickedInside Then
        
        If Not InvComNpc.IsItemSelected Then Exit Sub
    
        LasActionBuy = True

        If UserGLD >= InvComNpc.Valor(InvComNpc.SelectedItem) * Val(cantidad) Then
            Call WriteCommerceBuy(InvComNpc.SelectedItem, cantidad.Text)
        Else
            AddtoRichTextBox frmmain.RecTxt, "No tenés suficiente oro.", 2, 51, 223, 1, 1

        End If
        
    ElseIf InvComUsu.ClickedInside Then
    
        If Not InvComUsu.IsItemSelected Then Exit Sub
    
        ' Hacemos acción del doble clic correspondiente
        Dim ObjType As Byte

        ObjType = ObjData(InvComUsu.OBJIndex(InvComUsu.SelectedItem)).ObjType
        
        If UserMeditar Then Exit Sub
        If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
        
        Select Case ObjType

            Case eObjType.otArmadura, eObjType.otESCUDO, eObjType.otmagicos, eObjType.otFlechas, eObjType.otCASCO, eObjType.otNudillos
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

End Sub

Private Sub interface_KeyDown(KeyCode As Integer, Shift As Integer)

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

End Sub

Private Sub InvComUsu_ItemDropped(ByVal Drag As Integer, ByVal Drop As Integer, ByVal x As Integer, ByVal y As Integer)

    ' Si soltó dentro del mismo inventario
    If Drop > 0 Then
        ' Movemos el item dentro del inventario
        Call WriteItemMove(Drag, Drop)
    Else

        ' Si lo soltó dentro de la tienda
        If InvComNpc.GetSlot(x, y) > 0 Then
            ' Vendemos el item
            LasActionBuy = False
            Call WriteCommerceSell(Drag, max(Val(cantidad.Text), InvComUsu.Amount(InvComUsu.SelectedItem)))

        End If

    End If

End Sub

Private Sub InvComNpc_ItemDropped(ByVal Drag As Integer, ByVal Drop As Integer, ByVal x As Integer, ByVal y As Integer)

    ' Si lo soltó dentro del inventario
    If InvComUsu.GetSlot(x, y) > 0 Then
        ' Compramos el item
        LasActionBuy = True

        ' Si tiene suficiente oro
        If UserGLD >= InvComNpc.Valor(Drag) * Val(cantidad.Text) Then
            Call WriteCommerceBuy(Drag, Val(cantidad.Text))
        Else
            AddtoRichTextBox frmmain.RecTxt, "No tenés suficiente oro.", 2, 51, 223, 1, 1

        End If

    End If

End Sub

Private Sub salir_Click()
    Unload Me

End Sub

Private Sub tmrNumber_Timer()

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

End Sub
