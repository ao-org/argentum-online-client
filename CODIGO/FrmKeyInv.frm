VERSION 5.00
Begin VB.Form FrmKeyInv 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3600
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   188
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   240
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   600
      MousePointer    =   99  'Custom
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2400
   End
   Begin VB.Label NombreLlave 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Image cmdCerrar 
      Height          =   420
      Left            =   3135
      Tag             =   "0"
      Top             =   0
      Width           =   465
   End
End
Attribute VB_Name = "FrmKeyInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const WM_SYSCOMMAND As Long = &H112&
Const MOUSE_MOVE    As Long = &HF012&
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Public WithEvents InvKeys As clsGrapchicalInventory
Attribute InvKeys.VB_VarHelpID = -1

Private Sub cmdcerrar_Click()
    Unload Me
End Sub

Private Sub cmdCerrar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdCerrar.Picture = LoadInterface("boton-cerrar-off.bmp")
    cmdCerrar.Tag = "1"
End Sub

Private Sub imgCerrar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If cmdCerrar.Tag = "0" Then
        cmdCerrar.Picture = LoadInterface("boton-cerrar-over.bmp")
        cmdCerrar.Tag = "1"
    End If
End Sub

Private Sub Form_Load()
    Call FormParser.Parse_Form(Me)
    Me.Picture = LoadInterface("ventanallavero.bmp")
    
    If InvKeys.OBJIndex(0) = 0 Then
        NombreLlave.Caption = "Aquí aparecerán las llaves que consigas"
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    NombreLlave.Caption = vbNullString
    
    ReleaseCapture
    Call SendMessage(Me.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)
    
    If cmdCerrar.Tag = "1" Then
        cmdCerrar.Picture = Nothing
        cmdCerrar.Tag = "0"
    End If
End Sub

Private Sub interface_DblClick()
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub

    If InvKeys.IsItemSelected Then
        Call WriteUseKey(InvKeys.SelectedItem)
    End If
End Sub

Private Sub interface_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Slot As Integer
    Slot = InvKeys.GetSlot(x, y)
    
    If Slot <> 0 Then
        NombreLlave.Caption = InvKeys.ItemName(Slot)
    End If
    
    If cmdCerrar.Tag = "1" Then
        cmdCerrar.Picture = Nothing
        cmdCerrar.Tag = "0"
    End If
End Sub

Private Sub interface_Paint()
    InvKeys.ReDraw
End Sub
