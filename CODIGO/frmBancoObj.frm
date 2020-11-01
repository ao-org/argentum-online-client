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
      Height          =   4650
      Left            =   240
      MousePointer    =   99  'Custom
      ScaleHeight     =   310
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   507
      TabIndex        =   1
      Top             =   1650
      Width           =   7605
   End
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3630
      TabIndex        =   0
      Text            =   "1"
      Top             =   6645
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   510
      Index           =   1
      Left            =   5025
      Tag             =   "0"
      Top             =   6495
      Width           =   2220
   End
   Begin VB.Image Image1 
      Height          =   525
      Index           =   0
      Left            =   1125
      Tag             =   "0"
      Top             =   6495
      Width           =   1725
   End
End
Attribute VB_Name = "frmBancoObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const WM_SYSCOMMAND As Long = &H112&
Const MOUSE_MOVE As Long = &HF012&

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Long) As Long
Public LastIndex1 As Integer
Public LasActionBuy As Boolean

' Declaro los inventarios acá para manejar el evento drop
Public WithEvents InvBankUsu As clsGrapchicalInventory ' Inventario del usuario visible en la bóveda
Attribute InvBankUsu.VB_VarHelpID = -1
Public WithEvents InvBoveda As clsGrapchicalInventory ' Inventario de la bóveda
Attribute InvBoveda.VB_VarHelpID = -1

Private Sub moverForm()
    Dim res As Long
    ReleaseCapture
    res = SendMessage(Me.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)
End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If (KeyAscii = 27) Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
End Sub

Private Sub Image1_Click(Index As Integer)
Call Sound.Sound_Play(SND_CLICK)

If Not IsNumeric(cantidad.Text) Then Exit Sub

Select Case Index
    Case 0
        'frmBancoObj.List1(0).SetFocus
        'LastIndex1 = List1(0).ListIndex
        LasActionBuy = True
        'Call WriteBankExtractItem(InvBoveda.SelectedItem, cantidad.Text, 1)
        
        If InvBoveda.SelectedItem <= 0 Then Exit Sub
        Call WriteBankExtractItem(InvBoveda.SelectedItem, max(Val(cantidad.Text), InvBoveda.Amount(InvBoveda.SelectedItem)), 0)
   Case 1
        LasActionBuy = False
        If InvBankUsu.SelectedItem <= 0 Then Exit Sub
        Call WriteBankDeposit(InvBankUsu.SelectedItem, max(Val(cantidad.Text), InvBankUsu.Amount(InvBankUsu.SelectedItem)), 0)
End Select
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
End Sub
Private Sub cantidad_Change()
    If Val(cantidad.Text) < 1 Then
        cantidad.Text = 1
    End If
    
    If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
        cantidad.Text = MAX_INVENTORY_OBJS
        cantidad.SelStart = Len(cantidad.Text)
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call Sound.Sound_Play(SND_CLICK)
    Call WriteBankEnd
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Index = 0 Then
        'Image1(0).Picture = LoadInterface("retirarwidepress.bmp")
        Image1(0).Tag = "0"
Else
       ' Image1(1).Picture = LoadInterface("depowidepress.bmp")
        Image1(1).Tag = "0"
End If
End Sub


Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Index = 0 Then
    If Image1(0).Tag = "0" Then
        Image1(0).Picture = LoadInterface("retirarwidehover.bmp")
        Image1(0).Tag = "1"
    End If
Else
    
    If Image1(1).Tag = "0" Then
        Image1(1).Picture = LoadInterface("depowidehover.bmp")
        Image1(1).Tag = "1"
    End If
End If
End Sub

Private Sub interface_Click()
    
    If InvBoveda.ClickedInside Then
        ' Cliqueé en la bóveda, deselecciono el inventario
        Call InvBankUsu.SeleccionarItem(0)
        
    ElseIf InvBankUsu.ClickedInside Then
        ' Cliqueé en el inventario, deselecciono la bóveda
        Call InvBoveda.SeleccionarItem(0)
    End If

End Sub

Private Sub interface_DblClick()

    ' Nos aseguramos que lo último que cliqueó fue el inventario
    If Not InvBankUsu.ClickedInside Then Exit Sub

    ' Hacemos acción del doble clic correspondiente
    Dim ObjType As Byte
    ObjType = ObjData(InvBankUsu.OBJIndex(InvBankUsu.SelectedItem)).ObjType
    
    If Not IntervaloPermiteUsar Then Exit Sub
    
    Select Case ObjType
        Case eObjType.otArmadura, eObjType.otESCUDO, eObjType.OtHerramientas, eObjType.otmagicos, eObjType.otFlechas, eObjType.otCASCO, eObjType.otNudillos
            Call EquiparItem
            
        Case eObjType.otWeapon
            If ObjData(InvBankUsu.OBJIndex(InvBankUsu.SelectedItem)).proyectil = 1 And InvBankUsu.Equipped(InvBankUsu.SelectedItem) Then
                Call UsarItem
            ElseIf Not InvBankUsu.Equipped(InvBankUsu.SelectedItem) Then
                Call EquiparItem
            End If
             
        Case Else
            Call UsarItem
    End Select

End Sub

Private Sub interface_KeyDown(KeyCode As Integer, Shift As Integer)

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

End Sub

Private Sub InvBoveda_ItemDropped(ByVal Drag As Integer, ByVal Drop As Integer, ByVal x As Integer, ByVal y As Integer)

    ' Si lo soltó dentro de la bóveda
    If Drop > 0 Then
        ' Movemos el item dentro de la bóveda
        Call WriteBovedaItemMove(Drag, Drop)
    Else
        Drop = InvBankUsu.GetSlot(x, y)
        ' Si lo soltó dentro del inventario
        If Drop > 0 Then
            ' Retiramos el item
            Call WriteBankExtractItem(Drag, max(Val(cantidad.Text), InvBoveda.Amount(InvBoveda.SelectedItem)), Drop)
        End If
    End If

End Sub

Private Sub InvBankUsu_ItemDropped(ByVal Drag As Integer, ByVal Drop As Integer, ByVal x As Integer, ByVal y As Integer)

    ' Si lo soltó dentro del mismo inventario
    If Drop > 0 Then
        ' Movemos el item dentro del inventario
        Call WriteItemMove(Drag, Drop)
    Else
        Drop = InvBoveda.GetSlot(x, y)
        ' Si lo soltó dentro de la bóveda
        If Drop > 0 Then
            ' Depositamos el item
            Call WriteBankDeposit(Drag, max(Val(cantidad.Text), InvBankUsu.Amount(InvBankUsu.SelectedItem)), Drop)
        End If
    End If

End Sub
