VERSION 5.00
Begin VB.Form frmGuildAdm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Lista de clanes registrados"
   ClientHeight    =   5865
   ClientLeft      =   0
   ClientTop       =   -180
   ClientWidth     =   6225
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGuildAdm.frx":0000
   ScaleHeight     =   5865
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox GuildsList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2565
      ItemData        =   "frmGuildAdm.frx":768A4
      Left            =   495
      List            =   "frmGuildAdm.frx":768A6
      TabIndex        =   2
      Top             =   2160
      Width           =   4080
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   285
      ItemData        =   "frmGuildAdm.frx":768A8
      Left            =   2280
      List            =   "frmGuildAdm.frx":768B5
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1600
      Width           =   1655
   End
   Begin VB.TextBox qhi9t0 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2640
      TabIndex        =   0
      ToolTipText     =   "Nombre del clan a buscar"
      Top             =   1950
      Width           =   2115
   End
   Begin VB.Label lblClose 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   5760
      TabIndex        =   3
      Top             =   0
      Width           =   405
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   4645
      Tag             =   "0"
      Top             =   4230
      Width           =   390
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   480
      Tag             =   "0"
      Top             =   5040
      Width           =   1950
   End
   Begin VB.Image Image1 
      Height          =   425
      Left            =   4020
      Tag             =   "0"
      Top             =   1560
      Width           =   425
   End
End
Attribute VB_Name = "frmGuildAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Click()
    
    If Not ListaClanes Then Exit Sub

    frmGuildAdm.guildslist.Clear
    
    Dim i As Long

    For i = 0 To UBound(ClanesList)

        If Combo1.ListIndex < 2 Then
            If ClanesList(i).Alineacion = Combo1.ListIndex Then
                Call frmGuildAdm.guildslist.AddItem(ClanesList(i).nombre)

            End If

        Else
            
            Call frmGuildAdm.guildslist.AddItem(ClanesList(i).nombre)

        End If

    Next i

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then Unload Me

End Sub

Private Sub Form_Load()

    Call FormParser.Parse_Form(Me)
    
    Me.Picture = LoadInterface("VentanaClanes.bmp")
    
    
    Combo1.ListIndex = 2

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image1.Picture = Nothing
    Image1.Tag = "0"

    Image2.Picture = Nothing
    Image2.Tag = "0"

    Image3.Picture = Nothing
    Image3.Tag = "0"

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim b As Integer

    For b = 0 To guildslist.ListCount - 1
        guildslist.ListIndex = b

        If LCase$(guildslist) = LCase$(qhi9t0) Then
            Exit Sub

        End If

    Next
    MsgBox "Clan no encontrado"

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Image1.Tag = "0" Then
        Image1.Picture = LoadInterface("boton-buscar-over.bmp")
        Image1.Tag = "1"
    
    End If

End Sub

Private Sub Image2_Click()

    If UserEstado = 1 Then 'Muerto

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

        End With

        Exit Sub

    End If
                   
    Call WriteQuieroFundarClan

End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Image2.Tag = "0" Then
        Image2.Picture = LoadInterface("boton-fundar-clan-es-over.bmp")
        Image2.Tag = "1"

    End If

End Sub

Private Sub Image3_Click()
    
    'Si nos encontramos con un guild con nombre vacío algo sospechoso está pasando, x las dudas no hacemos nada.
    If Len(guildslist.List(guildslist.ListIndex)) = 0 Then Exit Sub
    
    frmGuildBrief.EsLeader = False
    
    Call WriteGuildRequestDetails(guildslist.List(guildslist.ListIndex))

End Sub


Private Sub lblClose_Click()
    Unload Me
End Sub
