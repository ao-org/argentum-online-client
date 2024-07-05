VERSION 5.00
Begin VB.Form frmGuildNews 
   BorderStyle     =   0  'None
   Caption         =   "Noticias de clan"
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   6555
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
   ScaleHeight     =   5640
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox miembros 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1590
      ItemData        =   "frmGuildNews.frx":0000
      Left            =   555
      List            =   "frmGuildNews.frx":0007
      TabIndex        =   5
      Top             =   3345
      Width           =   2340
   End
   Begin VB.ListBox guildslist 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1395
      ItemData        =   "frmGuildNews.frx":0015
      Left            =   3680
      List            =   "frmGuildNews.frx":0017
      TabIndex        =   4
      Top             =   3350
      Width           =   2325
   End
   Begin VB.TextBox news 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   3700
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   950
      Width           =   2260
   End
   Begin VB.Image cmdDetalles 
      Height          =   255
      Left            =   3840
      Top             =   4900
      Width           =   1815
   End
   Begin VB.Image cmdCerrar 
      Height          =   375
      Left            =   6090
      Top             =   10
      Width           =   375
   End
   Begin VB.Label lblNivel 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2455
      TabIndex        =   7
      Top             =   715
      Width           =   975
   End
   Begin VB.Label lblMiembros 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2215
      TabIndex        =   6
      Top             =   3130
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      Height          =   255
      Left            =   480
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label porciento 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
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
      Left            =   480
      TabIndex        =   2
      Top             =   1090
      Width           =   2415
   End
   Begin VB.Label beneficios 
      BackStyle       =   0  'Transparent
      Caption         =   "No atacarse / Chat de clan / Pedir ayuda (K) / Verse Invisible / Marca de clan / Verse vida / Max miembros: 25"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   1215
      Left            =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label expcount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "400 / 500"
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
      Left            =   480
      TabIndex        =   1
      Top             =   1090
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Shape EXPBAR 
      BackColor       =   &H00000080&
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   210
      Left            =   510
      Top             =   1095
      Width           =   960
   End
End
Attribute VB_Name = "frmGuildNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    Argentum 20 - Game Client Program
'    Copyright (C) 2022 - Noland Studios
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit
Private cBotonCerrar As clsGraphicalButton
Private cBotonDetalles As clsGraphicalButton
Private Sub beneficios_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo beneficios_MouseMove_Err
    
    porciento.Visible = True
    expcount.Visible = False

    
    Exit Sub

beneficios_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildNews.beneficios_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub loadButtons()
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonDetalles = New clsGraphicalButton
    
    Call cBotonCerrar.Initialize(cmdCerrar, "boton-cerrar-default.bmp", _
                                                "boton-cerrar-over.bmp", _
                                                "boton-cerrar-off.bmp", Me)
    Call cBotonDetalles.Initialize(cmdDetalles, "boton-detalles-default.bmp", _
                                                "boton-detalles-over.bmp", _
                                                "boton-detalles-off.bmp", Me)

End Sub

Private Sub cmdcerrar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    
    On Error GoTo Command1_Click_Err
    

    

    Unload Me
    frmMain.SetFocus

    
    Exit Sub

Command1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildNews.Command1_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdDetalles_Click()
    
    On Error GoTo cmdDetalles_Click_Err
    
    Call WriteGuildRequestDetails(guildslist.List(guildslist.ListIndex))

    
    Exit Sub

cmdDetalles_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmGuildNews.cmdDetalles_Click", Erl)
    Resume Next
    
End Sub

Private Sub Frame5_DragDrop(source As Control, x As Single, y As Single)
    
    On Error GoTo Frame5_DragDrop_Err
    
    porciento.Visible = True
    expcount.Visible = False

    
    Exit Sub

Frame5_DragDrop_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildNews.Frame5_DragDrop", Erl)
    Resume Next
    
End Sub


Private Sub porciento_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo porciento_MouseMove_Err
    
    porciento.Visible = False
    expcount.Visible = True

    
    Exit Sub

porciento_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildNews.porciento_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    Me.Picture = LoadInterface("ventanaclanes_noticias.bmp")
        
    Call Aplicar_Transparencia(Me.hwnd, 240)
    Call loadButtons
    Exit Sub
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildNews.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    Call MoverForm(Me.hwnd)
    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmGuildNews.Form_MouseMove", Erl)
    Resume Next
    
End Sub
