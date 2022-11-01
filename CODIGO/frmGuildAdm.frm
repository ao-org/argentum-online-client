VERSION 5.00
Begin VB.Form frmGuildAdm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Lista de clanes registrados"
   ClientHeight    =   5868
   ClientLeft      =   0
   ClientTop       =   -180
   ClientWidth     =   6228
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
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
   ScaleHeight     =   489
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   519
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Filtro 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   545
      TabIndex        =   2
      Top             =   1615
      Width           =   1575
   End
   Begin VB.ListBox GuildsList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2472
      ItemData        =   "frmGuildAdm.frx":0000
      Left            =   495
      List            =   "frmGuildAdm.frx":0002
      TabIndex        =   1
      Top             =   2160
      Width           =   4080
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      ItemData        =   "frmGuildAdm.frx":0004
      Left            =   2280
      List            =   "frmGuildAdm.frx":001A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1600
      Width           =   1655
   End
   Begin VB.Image cmdCerrar 
      Height          =   420
      Left            =   5760
      Top             =   0
      Width           =   420
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   4645
      Tag             =   "0"
      Top             =   4230
      Width           =   390
   End
   Begin VB.Image cmdFundarClan 
      Height          =   420
      Left            =   480
      Tag             =   "0"
      Top             =   5040
      Width           =   1950
   End
   Begin VB.Image cmdBuscar 
      Height          =   425
      Left            =   4005
      Tag             =   "0"
      Top             =   1560
      Width           =   450
   End
End
Attribute VB_Name = "frmGuildAdm"
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
Option Explicit

Private cBotonFundarClan As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton
Private cBotonBuscar As clsGraphicalButton


Private Sub cmdBuscar_Click()
    
    On Error GoTo Image1_Click_Err
    
    Dim i As Long

    frmGuildAdm.guildslist.Clear
    
    If Not ListaClanes Then Exit Sub
    
    If Len(Filtro.Text) <> 0 Then
        For i = 0 To UBound(ClanesList)

            If Combo1.ListIndex < 5 Then
                If ClanesList(i).Alineacion = Combo1.ListIndex Then
                    If InStr(1, UCase$(ClanesList(i).nombre), UCase$(Filtro.Text)) <> 0 Then
                        Call frmGuildAdm.guildslist.AddItem(ClanesList(i).nombre)
                    End If
                End If
            ElseIf InStr(1, UCase$(ClanesList(i).nombre), UCase$(Filtro.Text)) <> 0 Then
                Call frmGuildAdm.guildslist.AddItem(ClanesList(i).nombre)
            End If
    
        Next i
        
    Else
        For i = 0 To UBound(ClanesList)

            If Combo1.ListIndex < 5 Then
                If ClanesList(i).Alineacion = Combo1.ListIndex Then
                    Call frmGuildAdm.guildslist.AddItem(ClanesList(i).nombre)
    
                End If
    
            Else
                
                Call frmGuildAdm.guildslist.AddItem(ClanesList(i).nombre)
    
            End If
    
        Next i
    End If
    
    Exit Sub

Image1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildAdm.Image1_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdFundarClan_Click()
    On Error GoTo cmdFundarClan_Click_Err

    If UserEstado = 1 Then 'Muerto

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With

        Exit Sub

    End If
                   
    Call WriteQuieroFundarClan

    
    Exit Sub

cmdFundarClan_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmGuildAdm.cmdFundarClan_Click", Erl)
    Resume Next
End Sub


Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    

    Call FormParser.Parse_Form(Me)
    
    Me.Picture = LoadInterface("ventanaclanes.bmp")
    
    Call LoadButtons
    Combo1.ListIndex = 2

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmGuildAdm.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub LoadButtons()
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonFundarClan = New clsGraphicalButton
    Set cBotonBuscar = New clsGraphicalButton
    
    Call cBotonCerrar.Initialize(cmdCerrar, "boton-cerrar-default.bmp", _
                                                "boton-cerrar-over.bmp", _
                                                "boton-cerrar-off.bmp", Me)

    Call cBotonBuscar.Initialize(cmdBuscar, "boton-buscar-default.bmp", _
                                                    "boton-buscar-over.bmp", _
                                                    "boton-buscar-off.bmp", Me)
                                                    
    Call cBotonFundarClan.Initialize(cmdFundarClan, "boton-fundar-clan-default.bmp", _
                                                    "boton-fundar-clan-over.bmp", _
                                                    "boton-fundar-clan-off.bmp", Me)
End Sub

Private Sub Image3_Click()
    
    On Error GoTo Image3_Click_Err
    
    'Si nos encontramos con un guild con nombre vacío algo sospechoso está pasando, x las dudas no hacemos nada.
    If Len(guildslist.List(guildslist.ListIndex)) = 0 Then Exit Sub
    
    frmGuildBrief.EsLeader = False
    
    Call WriteGuildRequestDetails(guildslist.List(guildslist.ListIndex))
    
    Exit Sub

Image3_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildAdm.Image3_Click", Erl)
    Resume Next
    
End Sub


Private Sub lblClose_Click()
    Unload Me
End Sub
