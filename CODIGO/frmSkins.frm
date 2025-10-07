VERSION 5.00
Begin VB.Form frmSkins 
   BorderStyle     =   0  'None
   ClientHeight    =   7245
   ClientLeft      =   16350
   ClientTop       =   5160
   ClientWidth     =   3600
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   469.422
   ScaleMode       =   0  'User
   ScaleWidth      =   241
   ShowInTaskbar   =   0   'False
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
      Height          =   5820
      Left            =   240
      MousePointer    =   99  'Custom
      ScaleHeight     =   388
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   211
      TabIndex        =   0
      Top             =   720
      Width           =   3165
   End
   Begin VB.Image imgDeleteItem 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   240
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblItemData 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ItemData"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   6600
      Width           =   3420
   End
   Begin VB.Image cmdCerrar 
      Height          =   420
      Left            =   3120
      Tag             =   "0"
      Top             =   0
      Width           =   465
   End
End
Attribute VB_Name = "frmSkins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Constantes para simular el movimiento de la ventana
Private Const WM_SYSCOMMAND As Long = &H112&
Private Const MOUSE_MOVE    As Long = &HF012&
' Declaraciones de funciones API de Windows
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' Declaración de variable con evento de inventario gráfico
Private cBotonEliminarItem As clsGraphicalButton
Public WithEvents InvSkins As clsGrapchicalInventory
Attribute InvSkins.VB_VarHelpID = -1

' Evento al hacer clic en el botón "Cerrar"
Private Sub cmdCerrar_Click()
    ' Oculta el formulario
    'Call frmSkins.WalletSkins
    
    bSkins = False
    Unload Me
    
End Sub

Private Sub Form_Activate()
    Call InvSkins.ReDraw
End Sub

Private Sub Form_Initialize()
    
    bSkins = True

End Sub

Private Sub Form_Load()

    ' Parsea la interfaz del formulario (diseño)
   On Error GoTo Form_Load_Error

    Call FormParser.Parse_Form(Me)
    ' Aplica transparencia al formulario (valor 240 de opacidad)
    Call Aplicar_Transparencia(Me.hWnd, 240)
    ' Carga la imagen de fondo del formulario desde archivo
    frmSkins.Picture = LoadInterface("ventanaskins.bmp")
    
    Set cBotonEliminarItem = New clsGraphicalButton
    Call cBotonEliminarItem.Initialize(imgDeleteItem, "boton-borrar-item-default.bmp", "boton-borrar-item-over.bmp", "boton-borrar-item-off.bmp", Me)
    
    bSkins = True
    Call InvSkins.ReDraw
    DoEvents
    Exit Sub
    

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    Call RegistrarError(Err.Number, Err.Description, "frmSkins.Form_Load", Erl())
    
End Sub

' Permite mover el formulario arrastrándolo con el mouse
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = vbLeftButton Then
        Call ReleaseCapture
        Call SendMessage(Me.hWnd, WM_SYSCOMMAND, MOUSE_MOVE, 0&)
    End If

End Sub

' Método público para ocultar el formulario
Public Sub WalletSkins()
    
    On Error GoTo WalletSkins_Err
    Exit Sub
WalletSkins_Err:
    ' Manejo de errores estandarizado
    Call RegistrarError(Err.Number, Err.Description, "frmSkins.WalletSkins", Erl)
    Resume Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    bSkins = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bSkins = False
End Sub

Private Sub SkinEquip(ByVal eSkinType As eObjType)

Static lastItem                 As Long

Dim canEquip                    As Boolean

    If (Me.InvSkins.SelectedItem > 0) And (Me.InvSkins.SelectedItem < MAX_SKINSINVENTORY_SLOTS + 1) Then
        lastItem = Me.InvSkins.SelectedItem
        Call WriteEquipItem(Me.InvSkins.SelectedItem, True, eSkinType)
    End If
    
End Sub

Private Sub imgDeleteItem_Click()
    If Not InvSkins.IsItemSelected Then
        Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_NO_TIENE_ITEM_SELECCIONADO"), 255, 255, 255, False, False, False)
    Else
        If MsgBox(JsonLanguage.Item("MENSAJEBOX_ELIMINAR_ITEM"), vbYesNo, JsonLanguage.Item("MENSAJEBOX_TITULO_ELIMINAR_ITEM")) = vbYes Then
            Call WriteDeleteItem(InvSkins.SelectedItem, True)
        End If
    End If
End Sub

Private Sub interface_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case BindKeys(e_KeyAction.eEquipItem).KeyCode
            If Me.InvSkins.SelectedItem = 0 Then Exit Sub
            Call SkinEquip(Me.InvSkins.ObjType(Me.InvSkins.SelectedItem)) 'eObjType.otSkinsArmours)

        Case Else
            'do nothing?
    End Select
End Sub

Private Sub interface_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If InvSkins.SelectedItem > 0 Then
        Me.lblItemData.Caption = InvSkins.GetInfo(InvSkins.ObjIndex(InvSkins.SelectedItem))
    End If
End Sub

Private Sub interface_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call InvSkins.SeleccionarItem(0)
End Sub
