VERSION 5.00
Begin VB.Form FrmViajes 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6510
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
   ScaleHeight     =   5580
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   1290
      Left            =   360
      TabIndex        =   0
      Top             =   1630
      Width           =   2490
   End
   Begin VB.Image Image2 
      Height          =   2580
      Left            =   4320
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   3840
      Tag             =   "0"
      Top             =   4680
      Width           =   1890
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "FrmViajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)

    If (KeyAscii = 27) Then
        Unload Me

    End If

End Sub

Private Sub Form_Load()
    Call FormParser.Parse_Form(Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Image1.Tag = "1" Then
        Image1.Picture = Nothing
        Image2.Picture = Nothing
        Image1.Tag = "0"

    End If

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Image1.Tag = "0" Then
        Image1.Picture = LoadInterface("viajarhover" & ViajarInterface & ".bmp")
        Image2.Picture = LoadInterface("viaje" & ViajarInterface & "ok.bmp")
        Image1.Tag = "1"

    End If

End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim destino As Byte

    destino = List1.ListIndex + 1

    If destino <= 0 Then Exit Sub
    Unload Me
    Call WriteCompletarViaje(Destinos(destino).CityDest, Destinos(destino).costo)

End Sub
