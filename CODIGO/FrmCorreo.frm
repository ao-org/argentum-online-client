VERSION 5.00
Begin VB.Form FrmCorreo 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Correo de Argentum20"
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   8145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "FrmCorreo.frx":0000
   ScaleHeight     =   5850
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   3090
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   1995
   End
   Begin VB.CommandButton cmdClean 
      Appearance      =   0  'Flat
      Caption         =   "Limpiar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8520
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox ListaAenviar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   2370
      ItemData        =   "FrmCorreo.frx":4EDD8
      Left            =   5760
      List            =   "FrmCorreo.frx":4EDFA
      TabIndex        =   11
      Top             =   1700
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.TextBox txCantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3480
      TabIndex        =   3
      Text            =   "1"
      Top             =   4600
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.PictureBox picInvT 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   2800
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   10
      Top             =   4300
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txSndMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1620
      Left            =   450
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.TextBox txTo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   600
      TabIndex        =   7
      Top             =   1600
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.PictureBox picitem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   6400
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   3500
      Width           =   480
   End
   Begin VB.ListBox ListAdjuntos 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1650
      ItemData        =   "FrmCorreo.frx":4EE3A
      Left            =   5760
      List            =   "FrmCorreo.frx":4EE74
      TabIndex        =   2
      Top             =   1680
      Width           =   1770
   End
   Begin VB.TextBox txMensaje 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1995
      Left            =   3120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1680
      Width           =   2080
   End
   Begin VB.ListBox lstInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   2370
      ItemData        =   "FrmCorreo.frx":4EEE4
      Left            =   3000
      List            =   "FrmCorreo.frx":4EF1E
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "No adjuntaste nada:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6000
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image adjItem 
      Height          =   285
      Left            =   640
      Top             =   4260
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label lblCosto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   150
      Left            =   1440
      TabIndex        =   13
      Top             =   4640
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image Command3 
      Height          =   465
      Left            =   2100
      Tag             =   "0"
      Top             =   5260
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Image cmdSend 
      Height          =   465
      Left            =   5150
      Tag             =   "0"
      Top             =   5270
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Image Command1 
      Enabled         =   0   'False
      Height          =   480
      Left            =   5860
      Tag             =   "0"
      Top             =   4380
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Image Command2 
      Enabled         =   0   'False
      Height          =   495
      Left            =   4220
      Tag             =   "0"
      Top             =   4360
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Image Command9 
      Height          =   480
      Left            =   2880
      Tag             =   "0"
      Top             =   5250
      Width           =   2535
   End
   Begin VB.Image cmdSave 
      Height          =   465
      Left            =   4180
      Tag             =   "0"
      Top             =   4440
      Width           =   1485
   End
   Begin VB.Image cmdDel 
      Height          =   510
      Left            =   2720
      Tag             =   "0"
      Top             =   4410
      Width           =   1410
   End
   Begin VB.Label lbFecha 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   405
      Left            =   3120
      TabIndex        =   6
      Top             =   4080
      Width           =   1965
   End
   Begin VB.Label lbItem 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   540
      Left            =   5760
      TabIndex        =   5
      Top             =   4200
      Width           =   1725
   End
End
Attribute VB_Name = "FrmCorreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ItemRecibido(1 To 10) As Obj
Private Sub adjItem_Click()
lstInv.Enabled = Not lstInv.Enabled
txCantidad.Enabled = Not txCantidad.Enabled
ListaAenviar.Enabled = Not ListaAenviar.Enabled
Command1.Enabled = Not Command1.Enabled
Command2.Enabled = Not Command2.Enabled
If lstInv.Enabled Then
    adjItem.Picture = LoadInterface("check-amarillo.bmp")
    lblCosto.Caption = "Gratis"
Else
    lblCosto.Caption = "Gratis"
    adjItem.Picture = Nothing
End If
End Sub


Private Sub cmdClean_Click()
txTo.Text = ""
txSndMsg.Text = ""
'adjItem.value = 0
Dim i As Byte
    Call lstInv.Clear
    'Fill the inventory list
    For i = 1 To MAX_INVENTORY_SLOTS
        If frmmain.Inventario.OBJIndex(i) <> 0 Then
            lstInv.AddItem frmmain.Inventario.ItemName(i) & " - " & frmmain.Inventario.Amount(i)
        Else
            lstInv.AddItem "Vacio"
        End If
    Next i
    
    
    ListaAenviar.Clear

For i = 1 To 10
    ItemLista(i).OBJIndex = 0
    ItemLista(i).Amount = 0
    ListaAenviar.AddItem "Nada"
Next i


ItemCount = 0
Label4.Caption = "No agregaste ningun item"

End Sub

Private Sub cmdDel_Click()

If lstMsg.List(lstMsg.ListIndex) = "" Then Exit Sub

Call WriteBorrarCorreo(lstMsg.ListIndex + 1)
End Sub

Private Sub cmdDel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
           'cmdDel.Picture = LoadInterface("correo_borrarpress.bmp")
            'cmdDel.Tag = "1"
End Sub

Private Sub cmdDel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If cmdDel.Tag = "0" Then
        cmdDel.Picture = LoadInterface("correo_borrarhover.bmp")
        cmdDel.Tag = "1"
    End If
End Sub

Private Sub cmdSave_Click()
If lstMsg.ListIndex < 0 Then Exit Sub
Call WriteRetirarItemCorreo(lstMsg.ListIndex + 1)
End Sub

Private Sub cmdSave_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
                'cmdSave.Picture = LoadInterface("correo_guardaritempress.bmp")
                'cmdSave.Tag = "1"
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If cmdSave.Tag = "0" Then
        cmdSave.Picture = LoadInterface("correo_guardaritemhover.bmp")
        cmdSave.Tag = "1"
    End If
End Sub

Private Sub cmdSend_Click()
If txTo.Text = "" Then
    MsgBox ("¡Ingrese el nick del destinatario!")
    Exit Sub
End If




If adjItem Then
    If Not IsNumeric(txCantidad.Text) Or txCantidad.Text < 1 Or txCantidad.Text > 9999 Or txCantidad.Text > frmmain.Inventario.Amount(lstInv.ListIndex + 1) Then
        MsgBox ("¡Cantidad invalida!")
        Exit Sub
    End If
    If ItemCount <= 0 Then
        MsgBox ("¡Seleccione el item que desea enviar!")
        Exit Sub
    End If
    Call WriteSendCorreo(txTo.Text, txSndMsg.Text, ItemCount)
Else
    Call WriteSendCorreo(txTo.Text, txSndMsg.Text, 0)
End If
Unload Me
Call cmdClean_Click


End Sub

Private Sub cmdSend_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
                cmdSend.Picture = LoadInterface("correo_enviarpress.bmp")
                cmdSend.Tag = "1"
End Sub

Private Sub cmdSend_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If cmdSend.Tag = "0" Then
        cmdSend.Picture = LoadInterface("correo_enviarhover.bmp")
        cmdSend.Tag = "1"
    End If
End Sub

Private Sub Command1_Click()

If ItemCount = 0 Then
Label4.Caption = "No adjuntaste nada:"
End If


If ListaAenviar.ListIndex < 0 Then
    If ItemCount = 0 Then
        Label4.Caption = "No adjuntaste nada:"
    Else
        Label4.Caption = ItemCount & " items adjuntos:"
    End If
Exit Sub
End If
If ItemLista(ListaAenviar.ListIndex + 1).OBJIndex = 0 Then Exit Sub

ItemLista(ListaAenviar.ListIndex + 1).OBJIndex = 0
ListaAenviar.Clear

Dim i As Long
For i = 1 To 10

If ItemLista(i).OBJIndex = 0 Then
    If i = 10 Then
        ItemLista(i).OBJIndex = 0
        ItemLista(i).Amount = 0
        Exit For
    End If
        ItemLista(i).OBJIndex = ItemLista(i + 1).OBJIndex
        ItemLista(i).Amount = ItemLista(i + 1).Amount
        ItemLista(i + 1).OBJIndex = 0
        ItemLista(i + 1).Amount = 0
End If
Next i
    

For i = 1 To 10
    If ItemLista(i).OBJIndex > 0 Then
        ListaAenviar.AddItem frmmain.Inventario.ItemName(ItemLista(i).OBJIndex) & " - " & ItemLista(i).Amount
        ItemCount = i
    Else
        ListaAenviar.AddItem "Nada"
    End If
Next i

If ItemLista(1).OBJIndex = 0 Then
ItemCount = 0
Label4.Caption = "No agregaste ningun item"
Exit Sub
End If


Label4.Caption = ItemCount & " items adjuntos:"


End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        Command1.Picture = LoadInterface("correo_quitarpress.bmp")
        Command1.Tag = "1"
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Command1.Tag = "0" Then
        Command1.Picture = LoadInterface("correo_quitarhover.bmp")
        Command1.Tag = "1"
    End If
End Sub

Private Sub Command2_Click()

Dim i As Byte
Dim Encontre As Boolean
Dim Existia As Boolean
Dim NoTieneCantidad As Boolean

If frmmain.Inventario.Amount(lstInv.ListIndex + 1) = 0 Then Exit Sub

If frmmain.Inventario.Amount(lstInv.ListIndex + 1) < txCantidad.Text Then
    txCantidad.Text = frmmain.Inventario.Amount(lstInv.ListIndex + 1)
End If


ListaAenviar.Clear



For i = 1 To 10
    If ItemLista(i).OBJIndex = CByte(lstInv.ListIndex + 1) Then
        ItemLista(i).OBJIndex = CByte(lstInv.ListIndex + 1)
        ItemLista(i).Amount = ItemLista(i).Amount + CInt(txCantidad.Text)
        
        If ItemLista(i).Amount > frmmain.Inventario.Amount(ItemLista(i).OBJIndex) Then
            ItemLista(i).Amount = frmmain.Inventario.Amount(ItemLista(i).OBJIndex)
        End If
        
        Existia = True
        Encontre = True
        Exit For
    End If
Next i



If Existia = False Then
For i = 1 To 10
    If ItemLista(i).OBJIndex = 0 Then
        ItemLista(i).OBJIndex = CByte(lstInv.ListIndex + 1)
        ItemLista(i).Amount = CInt(txCantidad.Text)
        Encontre = True
        Exit For
    End If
Next i
End If



For i = 1 To 10
    If ItemLista(i).OBJIndex > 0 Then
        ListaAenviar.AddItem frmmain.Inventario.ItemName(ItemLista(i).OBJIndex) & " - " & ItemLista(i).Amount
        ItemCount = i
    Else
        ListaAenviar.AddItem "Nada"
    End If
Next i

Label4.Caption = ItemCount & " items adjuntos:"

ListaAenviar.Refresh

If Not Encontre Then MsgBox ("Solo podes enviar hasta 10 items.")

End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
                'Command2.Picture = LoadInterface("correo_agregarpress.bmp")
                'Command2.Tag = "1"
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Command2.Tag = "0" Then
        Command2.Picture = LoadInterface("correo_agregarhvoer.bmp")
        Command2.Tag = "1"
    End If
End Sub

Private Sub Command3_Click()
Me.Picture = LoadInterface("ventanacorreo.bmp")
lstMsg.Visible = True
txMensaje.Visible = True
ListAdjuntos.Visible = True
picitem.Visible = True
lbItem.Visible = True
lbFecha.Visible = True
cmdDel.Visible = True
cmdSave.Visible = True
Command9.Visible = True



txTo.Visible = False
txSndMsg.Visible = False
lstInv.Visible = False
txCantidad.Visible = False
ListaAenviar.Visible = False
picInvT.Visible = False
lblCosto.Visible = False
Command2.Visible = False
Command1.Visible = False
cmdSend.Visible = False
Command3.Visible = False
adjItem.Visible = False
Label4.Visible = False
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
               ' Command3.Picture = LoadInterface("correo_atraspress.bmp")
               ' Command3.Tag = "1"
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Command3.Tag = "0" Then
        Command3.Picture = LoadInterface("correo_atrashover.bmp")
        Command3.Tag = "1"
    End If
End Sub

Private Sub Command9_Click()
lstMsg.Visible = False
txMensaje.Visible = False
ListAdjuntos.Visible = False
picitem.Visible = False
lbItem.Visible = False
lbFecha.Visible = False
cmdDel.Visible = False
cmdSave.Visible = False
Command9.Visible = False
txTo.Visible = True
txSndMsg.Visible = True
lstInv.Visible = True
txCantidad.Visible = True
ListaAenviar.Visible = True
picInvT.Visible = True
lblCosto.Visible = True
Command2.Visible = True
Command1.Visible = True
cmdSend.Visible = True
Command3.Visible = True
adjItem.Visible = True
Label4.Visible = True
Me.Picture = LoadInterface("ventananuevocorreo.bmp")
End Sub

Private Sub Command9_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
                'Command9.Picture = LoadInterface("correo_nuevomensajepress.bmp")
                'command9.Tag = "1"
End Sub
Private Sub Command9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Command9.Tag = "0" Then
        Command9.Picture = LoadInterface("correo_nuevomensajehover.bmp")
        Command9.Tag = "1"
    End If
End Sub




Private Sub Form_KeyPress(KeyAscii As Integer)
If (KeyAscii = 27) Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
        Dim i As Byte
        For i = 1 To 10
          ItemLista(i).OBJIndex = 0
          ItemLista(i).Amount = 0
        Next i
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command9.Picture = Nothing
Command9.Tag = "0"

cmdDel.Picture = Nothing
cmdDel.Tag = "0"

cmdSave.Picture = Nothing
cmdSave.Tag = "0"

Command3.Picture = Nothing
Command3.Tag = "0"


cmdSend.Picture = Nothing
cmdSend.Tag = "0"

Command1.Picture = Nothing
Command1.Tag = "0"

Command2.Picture = Nothing
Command2.Tag = "0"

Command2.Picture = Nothing
Command2.Tag = "0"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Byte
    For i = 1 To 10
         ItemLista(i).OBJIndex = 0
         ItemLista(i).Amount = 0
    Next i
End Sub

Private Sub Image1_Click()
lstInv.Enabled = Not lstInv.Enabled
txCantidad.Enabled = Not txCantidad.Enabled
ListaAenviar.Enabled = Not ListaAenviar.Enabled
Command1.Enabled = Not Command1.Enabled
Command2.Enabled = Not Command2.Enabled
If lstInv.Enabled Then
    lblCosto.Caption = "Gratis"
Else
    lblCosto.Caption = "Gratis"
End If
End Sub



Private Sub ListaAenviar_Click()
If ListaAenviar.ListIndex + 1 > 10 Then Exit Sub
If ItemLista(ListaAenviar.ListIndex + 1).OBJIndex = 0 Then
picInvT.BackColor = vbBlack
Else
Call Grh_Render_To_Hdc(picInvT, frmmain.Inventario.GrhIndex(ItemLista(ListaAenviar.ListIndex + 1).OBJIndex), 0, 0)
End If
End Sub

Private Sub ListaAenviar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command9.Picture = Nothing
Command9.Tag = "0"

cmdDel.Picture = Nothing
cmdDel.Tag = "0"

cmdSave.Picture = Nothing
cmdSave.Tag = "0"

Command3.Picture = Nothing
Command3.Tag = "0"


cmdSend.Picture = Nothing
cmdSend.Tag = "0"

Command1.Picture = Nothing
Command1.Tag = "0"

Command2.Picture = Nothing
Command2.Tag = "0"

Command2.Picture = Nothing
Command2.Tag = "0"
End Sub

Private Sub ListAdjuntos_Click()
    picitem.BackColor = RGB(26, 16, 8)
    FrmCorreo.picitem.Refresh
    Call Grh_Render_To_Hdc(picitem, ObjData(ItemRecibido(ListAdjuntos.ListIndex + 1).OBJIndex).GrhIndex, 0, 0)
    lbItem.Caption = ObjData(ItemRecibido(ListAdjuntos.ListIndex + 1).OBJIndex).name & " (" & ItemRecibido(ListAdjuntos.ListIndex + 1).Amount & ")"
End Sub

Private Sub ListAdjuntos_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command9.Picture = Nothing
Command9.Tag = "0"

cmdDel.Picture = Nothing
cmdDel.Tag = "0"

cmdSave.Picture = Nothing
cmdSave.Tag = "0"

Command3.Picture = Nothing
Command3.Tag = "0"


cmdSend.Picture = Nothing
cmdSend.Tag = "0"

Command1.Picture = Nothing
Command1.Tag = "0"

Command2.Picture = Nothing
Command2.Tag = "0"

Command2.Picture = Nothing
Command2.Tag = "0"
End Sub

Private Sub lstInv_Click()

If frmmain.Inventario.GrhIndex(lstInv.ListIndex + 1) = 0 Then
picInvT.BackColor = vbBlack
Else

Call Grh_Render_To_Hdc(picInvT, frmmain.Inventario.GrhIndex(lstInv.ListIndex + 1), 0, 0)
End If
End Sub

Private Sub lstInv_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command9.Picture = Nothing
Command9.Tag = "0"

cmdDel.Picture = Nothing
cmdDel.Tag = "0"

cmdSave.Picture = Nothing
cmdSave.Tag = "0"

Command3.Picture = Nothing
Command3.Tag = "0"


cmdSend.Picture = Nothing
cmdSend.Tag = "0"

Command1.Picture = Nothing
Command1.Tag = "0"

Command2.Picture = Nothing
Command2.Tag = "0"

Command2.Picture = Nothing
Command2.Tag = "0"
End Sub

Private Sub lstMsg_Click()
Dim rdata As String
Dim item As String
Dim Index As Long
Dim cantidad As String

ListAdjuntos.Clear
If lstMsg.ListIndex < 0 Then Exit Sub
txMensaje = CorreoMsj(lstMsg.ListIndex + 1).mensaje
If CorreoMsj(lstMsg.ListIndex + 1).ItemCount > 0 Then
    'FrmCorreo.picItem.Visible = True
    cmdSave.Enabled = True
    
    
    
    Dim i As Long
    For i = 1 To CorreoMsj(lstMsg.ListIndex + 1).ItemCount
    
        rdata = Right$(CorreoMsj(lstMsg.ListIndex + 1).ItemArray, Len(CorreoMsj(lstMsg.ListIndex + 1).ItemArray))
        item = (ReadField(i, rdata, Asc("@")))
        
                
        rdata = Left$(item, Len(item))
        Index = (ReadField(1, rdata, Asc("-")))
        
        rdata = Right$(item, Len(item))
        cantidad = (ReadField(2, rdata, Asc("-")))
        
        ItemRecibido(i).OBJIndex = Index
        ItemRecibido(i).Amount = cantidad
        ListAdjuntos.AddItem ObjData(Index).name
    Next i
    
    ListAdjuntos.Enabled = True

   ' For i = 1 To CorreoMsj(lstMsg.ListIndex + 1).ItemCount
    
    'FrmCorreo.picItem.Refresh
   ' Call Grh_Render_To_Hdc(picItem, ObjData(CorreoMsj(lstMsg.ListIndex + 1).item.OBJIndex).GrhIndex, 0, 0)
    'lbItem.Caption = ObjData(CorreoMsj(lstMsg.ListIndex + 1).item.OBJIndex).name & " (" & CorreoMsj(lstMsg.ListIndex + 1).item.Amount & ")"
Else

ListAdjuntos.Enabled = False

    picitem.BackColor = RGB(0, 0, 0)
    FrmCorreo.picitem.Refresh
   
    'FrmCorreo.picItem.Visible = False
    cmdSave.Enabled = False
    lbItem.Caption = ""
   
End If

lbFecha.Caption = "Fecha de envío: " & CorreoMsj(lstMsg.ListIndex + 1).Fecha
End Sub

Private Sub txMensaje_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command9.Picture = Nothing
Command9.Tag = "0"

cmdDel.Picture = Nothing
cmdDel.Tag = "0"

cmdSave.Picture = Nothing
cmdSave.Tag = "0"

Command3.Picture = Nothing
Command3.Tag = "0"


cmdSend.Picture = Nothing
cmdSend.Tag = "0"

Command1.Picture = Nothing
Command1.Tag = "0"

Command2.Picture = Nothing
Command2.Tag = "0"

Command2.Picture = Nothing
Command2.Tag = "0"
End Sub

