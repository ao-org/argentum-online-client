VERSION 5.00
Begin VB.Form frmGuildLeader 
   BorderStyle     =   0  'None
   Caption         =   "Administración del Clan"
   ClientHeight    =   8376
   ClientLeft      =   0
   ClientTop       =   -180
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   698
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtguildnews 
      BackColor       =   &H00070707&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1170
      Left            =   330
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   6240
      Width           =   7560
   End
   Begin VB.ListBox solicitudes 
      Appearance      =   0  'Flat
      BackColor       =   &H00070707&
      ForeColor       =   &H00FFFFFF&
      Height          =   1044
      ItemData        =   "frmGuildLeader.frx":0000
      Left            =   4185
      List            =   "frmGuildLeader.frx":0002
      TabIndex        =   2
      Top             =   1590
      Width           =   3735
   End
   Begin VB.ListBox guildslist 
      Appearance      =   0  'Flat
      BackColor       =   &H00070707&
      ForeColor       =   &H00FFFFFF&
      Height          =   1248
      ItemData        =   "frmGuildLeader.frx":0004
      Left            =   4200
      List            =   "frmGuildLeader.frx":0006
      TabIndex        =   1
      Top             =   3720
      Width           =   3735
   End
   Begin VB.ListBox members 
      Appearance      =   0  'Flat
      BackColor       =   &H00070707&
      ForeColor       =   &H00FFFFFF&
      Height          =   1248
      ItemData        =   "frmGuildLeader.frx":0008
      Left            =   315
      List            =   "frmGuildLeader.frx":000A
      TabIndex        =   0
      Top             =   3720
      Width           =   3735
   End
   Begin VB.Label nivel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Alegreya Sans AO"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000CB1FE&
      Height          =   240
      Left            =   870
      TabIndex        =   9
      Top             =   1575
      Width           =   195
   End
   Begin VB.Label expcount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "400 / 500"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   540
      TabIndex        =   6
      Top             =   1965
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.Label maxMiembros 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "Alegreya Sans AO"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000CB1FE&
      Height          =   240
      Left            =   3525
      TabIndex        =   8
      Top             =   2295
      Width           =   210
   End
   Begin VB.Image cmdCerrar 
      Height          =   420
      Left            =   7815
      Top             =   0
      Width           =   465
   End
   Begin VB.Image cmdActualizar 
      Height          =   420
      Left            =   1260
      Top             =   7725
      Width           =   1980
   End
   Begin VB.Image cmdEditarDescripcion 
      Height          =   420
      Left            =   3780
      Top             =   7725
      Width           =   3240
   End
   Begin VB.Label porciento 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   540
      TabIndex        =   5
      Top             =   1965
      Width           =   3285
   End
   Begin VB.Label Miembros 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Alegreya Sans AO"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000CB1FE&
      Height          =   240
      Left            =   2100
      TabIndex        =   4
      Top             =   2280
      Width           =   195
   End
   Begin VB.Label beneficios 
      BackStyle       =   0  'Transparent
      Caption         =   "No atacarse / Chat de clan / Pedir ayuda (K)  / Verse Invisible / Marca de clan / Verse vida."
      BeginProperty Font 
         Name            =   "Alegreya Sans AO"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000CB1FE&
      Height          =   735
      Left            =   390
      TabIndex        =   3
      Top             =   2775
      Width           =   3615
   End
   Begin VB.Image EXPBAR 
      Height          =   225
      Left            =   390
      Top             =   1950
      Width           =   3585
   End
End
Attribute VB_Name = "frmGuildLeader"
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
Private cBotonActualizar As clsGraphicalButton
Private cBotonEditarDescripcion As clsGraphicalButton

Private Sub cmdElecciones_Click()
    
    On Error GoTo cmdElecciones_Click_Err
    
    Call WriteGuildOpenElections
    Unload Me

    
    Exit Sub

cmdElecciones_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildLeader.cmdElecciones_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command1_Click()
    
    On Error GoTo Command1_Click_Err
    Exit Sub

    If solicitudes.ListIndex = -1 Then Exit Sub
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembershipRequests
    'Call WriteGuildMemberInfo(solicitudes.List(solicitudes.ListIndex))

    'Unload Me
    
    Exit Sub

Command1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildLeader.Command1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command2_Click()
    
    On Error GoTo Command2_Click_Err
    

    If members.ListIndex = -1 Then Exit Sub
    
    frmCharInfo.frmType = CharInfoFrmType.frmMembers
    Call WriteGuildMemberInfo(members.List(members.ListIndex))

    'Unload Me
    
    Exit Sub

Command2_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildLeader.Command2_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdActualizar_Click()
    Dim k As String

    k = Replace(txtguildnews, vbCrLf, "º")
    
    Call WriteGuildUpdateNews(k)

    
    Exit Sub
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdEditarDescripcion_Click()
    Dim fdesc As String

    fdesc = InputBox("Ingrese la descripción:", "Modificar descripción")

    fdesc = Replace(fdesc, vbCrLf, "º", , , vbBinaryCompare)
    
    If Not ValidDescriptionCharacters(fdesc) Then
        MsgBox "La descripcion contiene caracteres invalidos"
        Exit Sub
    Else
        Call WriteClanCodexUpdate(fdesc)

    End If

    
    Exit Sub

End Sub

Private Sub Command4_Click()
    
    On Error GoTo Command4_Click_Err
    
    frmGuildBrief.EsLeader = True
    Call WriteGuildRequestDetails(guildslist.List(guildslist.ListIndex))

    'Unload Me
    
    Exit Sub

Command4_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmGuildLeader.Command4_Click", Erl)
    Resume Next
    
End Sub


Private Sub Command6_Click()

    'Call frmGuildURL.Show(vbModeless, frmGuildLeader)
    'Unload Me
End Sub

Private Sub Command7_Click()
    
    On Error GoTo Command7_Click_Err
    
    Call WriteGuildPeacePropList

    
    Exit Sub

Command7_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildLeader.Command7_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command9_Click()
    
    On Error GoTo Command9_Click_Err
    
    Call WriteGuildAlliancePropList

    
    Exit Sub

Command9_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildLeader.Command9_Click", Erl)
    Resume Next
    
End Sub

Private Sub expne_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo expne_MouseMove_Err
    
    porciento.Visible = True
    expcount.Visible = False

    
    Exit Sub

expne_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildLeader.expne_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
   
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    Me.Picture = LoadInterface("ventanaadminclan.bmp")
    EXPBAR.Picture = LoadInterface("barra-nivel-clan.bmp")
    
    Call Aplicar_Transparencia(Me.hwnd, 240)
    
    Call LoadButtons
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildLeader.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub LoadButtons()
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonActualizar = New clsGraphicalButton
    Set cBotonEditarDescripcion = New clsGraphicalButton
    
    Call cBotonCerrar.Initialize(cmdCerrar, "boton-cerrar-default.bmp", _
                                                "boton-cerrar-over.bmp", _
                                                "boton-cerrar-off.bmp", Me)
                                                
    Call cBotonActualizar.Initialize(cmdActualizar, "boton-actualizar-default.bmp", _
                                                "boton-actualizar-over.bmp", _
                                                "boton-actualizar-off.bmp", Me)
    Call cBotonEditarDescripcion.Initialize(cmdEditarDescripcion, "boton-editar-desc-clan-default.bmp", _
                                                "boton-editar-desc-clan-over.bmp", _
                                                "boton-editar-desc-clan-off.bmp", Me)
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    Call MoverForm(Me.hwnd)
    
    porciento.Visible = True
    expcount.Visible = False

    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildLeader.Form_MouseMove", Erl)
    Resume Next
    
End Sub


Private Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Frame4_MouseMove_Err
    
    porciento.Visible = True
    expcount.Visible = False

    
    Exit Sub

Frame4_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildLeader.Frame4_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub guildslist_DblClick()
   If guildslist.ListIndex > -1 Then
        frmGuildBrief.EsLeader = True
        Call WriteGuildRequestDetails(guildslist.List(guildslist.ListIndex))
        Exit Sub
    End If
End Sub

Private Sub members_DblClick()
  If members.ListIndex = -1 Then Exit Sub
    frmCharInfo.frmType = CharInfoFrmType.frmMembers
    Call WriteGuildMemberInfo(members.List(members.ListIndex))
End Sub

Private Sub solicitudes_DblClick()
 If solicitudes.ListIndex = -1 Then Exit Sub
    frmCharInfo.frmType = CharInfoFrmType.frmMembershipRequests
    Call WriteGuildMemberInfo(solicitudes.List(solicitudes.ListIndex))
    Exit Sub
End Sub

Private Sub porciento_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo porciento_MouseMove_Err
    
    porciento.Visible = False
    expcount.Visible = True

    
    Exit Sub

porciento_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildLeader.porciento_MouseMove", Erl)
    Resume Next
    
End Sub

