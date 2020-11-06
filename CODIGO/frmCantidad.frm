VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1860
   ClientLeft      =   1635
   ClientTop       =   4410
   ClientWidth     =   3240
   ControlBox      =   0   'False
   FillColor       =   &H00C00000&
   ForeColor       =   &H8000000D&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   3240
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   225
      Left            =   960
      MaxLength       =   9999
      TabIndex        =   0
      Text            =   "1"
      Top             =   940
      Width           =   1320
   End
   Begin VB.Image tirartodo 
      Height          =   420
      Left            =   1400
      Tag             =   "1"
      Top             =   1360
      Width           =   1740
   End
   Begin VB.Image tirar 
      Height          =   420
      Left            =   90
      Tag             =   "1"
      Top             =   1360
      Width           =   1305
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
Public bmoving As Boolean
Public dX As Integer
Public dy As Integer
Option Explicit
' Constantes para SendMessage
Const WM_SYSCOMMAND As Long = &H112&
Const MOUSE_MOVE As Long = &HF012&

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Long) As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
If (KeyAscii = 27) Then
    Unload Me
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub moverForm()
    Dim res As Long
    ReleaseCapture
    res = SendMessage(Me.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)
End Sub
Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
Text1.SelStart = 1
End Sub
Private Sub Text1_Change()

On Error GoTo errhandler
    If Val(Text1.Text) < 0 Then
        Text1.Text = "1"
    End If
    
    If Val(Text1.Text) > MAX_INVENTORY_OBJS Then
        Text1.Text = "10000"
        Text1.SelStart = Len(Text1.Text)
    End If
    
    Exit Sub
errhandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    Text1.Text = "1"
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
moverForm
If tirar.Tag = "1" Then
    tirar.Picture = Nothing
    tirar.Tag = "0"
End If

If tirartodo.Tag = "1" Then
    tirartodo.Picture = Nothing
    tirartodo.Tag = "0"
End If


End Sub
Private Sub tirar_click()
    
    If Not MainTimer.Check(TimersIndex.Drop) Then Exit Sub

Call Sound.Sound_Play(SND_CLICK)
    If LenB(frmCantidad.Text1.Text) > 0 Then
        If Not IsNumeric(frmCantidad.Text1.Text) Then Exit Sub  'Should never happen
        
        If frmmain.Inventario.SelectedItem <> FLAGORO Then
        
            If ObjData(frmmain.Inventario.OBJIndex(frmmain.Inventario.SelectedItem)).Destruye = 0 Then
                Call WriteDrop(frmmain.Inventario.SelectedItem, frmCantidad.Text1.Text)
            Else
                PreguntaScreen = "El item se destruira al tirarlo ¿Esta seguro?"
                Pregunta = True
                DestItemSlot = frmmain.Inventario.SelectedItem
                DestItemCant = frmCantidad.Text1.Text
                
                PreguntaLocal = True
                PreguntaNUM = 1
            End If
        Else
            Call WriteDrop(frmmain.Inventario.SelectedItem, frmCantidad.Text1.Text)
        End If
        
        frmCantidad.Text1.Text = ""
    End If
    Unload Me
End Sub
Private Sub tirar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        'tirar.Picture = LoadInterface("botontirarapretado.bmp")
End Sub
Private Sub tirar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If tirar.Tag = "0" Then
        tirar.Picture = LoadInterface("botontirarmarcado.bmp")
        tirar.Tag = "1"
    End If
End Sub


Private Sub tirartodo_click()
    If Not MainTimer.Check(TimersIndex.Drop) Then Exit Sub

    Call Sound.Sound_Play(SND_CLICK)
    If frmmain.Inventario.SelectedItem = 0 Then Exit Sub
    

    If frmmain.Inventario.SelectedItem <> FLAGORO Then
        If ObjData(frmmain.Inventario.OBJIndex(frmmain.Inventario.SelectedItem)).Destruye = 0 Then
            Call WriteDrop(frmmain.Inventario.SelectedItem, frmmain.Inventario.Amount(frmmain.Inventario.SelectedItem))
        Else
            PreguntaScreen = "El item se destruira al tirarlo ¿Esta seguro?"
            Pregunta = True
            DestItemSlot = frmmain.Inventario.SelectedItem
            DestItemCant = frmmain.Inventario.Amount(frmmain.Inventario.SelectedItem)
            
            PreguntaLocal = True
            PreguntaNUM = 1
        End If
        Unload Me
    Else
        If UserGLD > 100000 Then
            Call WriteDrop(frmmain.Inventario.SelectedItem, 100000)
            Unload Me
        Else
            Call WriteDrop(frmmain.Inventario.SelectedItem, UserGLD)
            Unload Me
        End If
    End If

    frmCantidad.Text1.Text = ""
End Sub
Private Sub tirartodo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        'tirartodo.Picture = LoadInterface("botontirartodoapretado.bmp")
End Sub
Private Sub tirartodo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If tirartodo.Tag = "0" Then
        tirartodo.Picture = LoadInterface("botontirartodomarcado.bmp")
        tirartodo.Tag = "1"
    End If
End Sub

