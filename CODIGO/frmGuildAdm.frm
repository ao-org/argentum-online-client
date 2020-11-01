VERSION 5.00
Begin VB.Form frmGuildAdm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Lista de clanes registrados"
   ClientHeight    =   6675
   ClientLeft      =   0
   ClientTop       =   -180
   ClientWidth     =   5595
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox GuildsList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      ItemData        =   "frmGuildAdm.frx":0000
      Left            =   840
      List            =   "frmGuildAdm.frx":0002
      TabIndex        =   2
      Top             =   3390
      Width           =   3830
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
      ItemData        =   "frmGuildAdm.frx":0004
      Left            =   2560
      List            =   "frmGuildAdm.frx":0011
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2580
      Width           =   2250
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
   Begin VB.Image Image3 
      Height          =   555
      Left            =   2760
      Tag             =   "0"
      Top             =   5730
      Width           =   2280
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   770
      Tag             =   "0"
      Top             =   5780
      Width           =   1890
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   4940
      Tag             =   "0"
      Top             =   1920
      Width           =   270
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

    frmGuildAdm.GuildsList.Clear
    
    Dim i As Long
    For i = 0 To UBound(ClanesList)
        If Combo1.ListIndex < 2 Then
            If ClanesList(i).Alineacion = Combo1.ListIndex Then
                Call frmGuildAdm.GuildsList.AddItem(ClanesList(i).nombre)
            End If
        Else
            
            Call frmGuildAdm.GuildsList.AddItem(ClanesList(i).nombre)
        End If
    Next i
End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
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
For b = 0 To GuildsList.ListCount - 1
    GuildsList.ListIndex = b
    If LCase$(GuildsList) = LCase$(qhi9t0) Then
        Exit Sub
    End If
Next
MsgBox "Clan no encontrado"
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Image1.Tag = "0" Then
        Image1.Picture = LoadInterface("clan_buscarclan.bmp")
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
        Image2.Picture = LoadInterface("clan_fundarapretado.bmp")
        Image2.Tag = "1"
    End If
End Sub

Private Sub Image3_Click()
    frmGuildBrief.EsLeader = False
    Call WriteGuildRequestDetails(GuildsList.List(GuildsList.ListIndex))
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Image3.Tag = "0" Then
        Image3.Picture = LoadInterface("clan_detallesclanapretado.bmp")
        Image3.Tag = "1"
    End If
End Sub

