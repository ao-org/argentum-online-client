VERSION 5.00
Begin VB.Form frmCarp 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Trabajar de carpintero"
   ClientHeight    =   6528
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7068
   BeginProperty Font 
      Name            =   "Verdana"
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
   ScaleHeight     =   544
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   589
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstArmas 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3696
      Left            =   525
      TabIndex        =   4
      Top             =   1440
      Width           =   2700
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1248
      Left            =   3840
      TabIndex        =   3
      Top             =   2520
      Width           =   1845
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1248
      ItemData        =   "frmCarp.frx":0000
      Left            =   5760
      List            =   "frmCarp.frx":0007
      TabIndex        =   2
      Top             =   2520
      Width           =   645
   End
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000D1213&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   250
      Left            =   5205
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "1"
      Top             =   4245
      Width           =   660
   End
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   4890
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   1845
      Width           =   480
   End
   Begin VB.Image cmdMas 
      Height          =   315
      Left            =   6000
      Tag             =   "0"
      Top             =   4215
      Width           =   315
   End
   Begin VB.Image cmdMenos 
      Height          =   315
      Left            =   4740
      Tag             =   "0"
      Top             =   4215
      Width           =   315
   End
   Begin VB.Image cmdCerrar 
      Height          =   420
      Left            =   6600
      Tag             =   "0"
      Top             =   0
      Width           =   420
   End
   Begin VB.Label desc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   3930
      Width           =   2775
   End
   Begin VB.Image cmdAceptar 
      Height          =   420
      Left            =   2520
      Tag             =   "0"
      Top             =   5730
      Width           =   1980
   End
   Begin VB.Image cmdConstruir 
      Height          =   420
      Left            =   4155
      Tag             =   "0"
      Top             =   4680
      Width           =   1980
   End
End
Attribute VB_Name = "frmCarp"
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


Private cBotonAceptar As clsGraphicalButton
Private cBotonConstruir As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton
Private cBotonMas As clsGraphicalButton
Private cBotonMenos As clsGraphicalButton



Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call Aplicar_Transparencia(Me.hwnd, 240)
    
    Call FormParser.Parse_Form(Me)
    
    Me.Picture = LoadInterface("VentanaCarpinteria.bmp")
    
    Call LoadButtons
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCarp.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub LoadButtons()
       
    Set cBotonAceptar = New clsGraphicalButton
    Set cBotonConstruir = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonMas = New clsGraphicalButton
    Set cBotonMenos = New clsGraphicalButton


    Call cBotonAceptar.Initialize(cmdAceptar, "boton-aceptar-default.bmp", _
                                                "boton-aceptar-over.bmp", _
                                                "boton-aceptar-off.bmp", Me)
    
    Call cBotonConstruir.Initialize(cmdConstruir, "boton-construir-default.bmp", _
                                                "boton-construir-over.bmp", _
                                                "boton-construir-off.bmp", Me)
                                                
    Call cBotonCerrar.Initialize(cmdCerrar, "boton-cerrar-default.bmp", _
                                                "boton-cerrar-over.bmp", _
                                                "boton-cerrar-off.bmp", Me)
                                                
    Call cBotonMas.Initialize(cmdMas, "boton-sm-mas-default.bmp", _
                                                "boton-sm-mas-over.bmp", _
                                                "boton-sm-mas-off.bmp", Me)
                                                
    Call cBotonMenos.Initialize(cmdMenos, "boton-sm-menos-default.bmp", _
                                                "boton-sm-menos-over.bmp", _
                                                "boton-sm-menos-off.bmp", Me)
End Sub

Private Sub cmdcerrar_Click()
    Unload Me
End Sub

Private Sub cmdMenos_Click()
    If cantidad > 0 Then
        cantidad = cantidad - 1
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdMas_Click()
    If cantidad <= 9999 Then
        cantidad = cantidad + 1
    Else
        Exit Sub
    End If
End Sub
Private Sub cmdConstruir_Click()
    
    On Error GoTo cmdConstruir_Click_Err
    
    
    'Si el indice seleccionado es -1 es xq no seleccionamos un item de la lista.
    If lstArmas.ListIndex = -1 Then Exit Sub

    If cantidad > 0 Then
        Call WriteCraftCarpenter(ObjCarpintero(lstArmas.ListIndex + 1), CLng(cantidad))
        Call AddtoRichTextBox(frmMain.RecTxt, "Comienzas a trabajar.", 2, 51, 223, 1, 1)
    Else
        Call AddtoRichTextBox(frmMain.RecTxt, "La cantidad debe ser mayor a 0.", 2, 51, 223, 1, 1)
    End If

    Unload Me
    
    
    Exit Sub

cmdConstruir_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCarp.cmdConstruir_Click", Erl)
    Resume Next
    
End Sub



Private Sub cmdAceptar_Click()
    
    On Error GoTo cmdAceptar_Click_Err
    
    Unload Me
    
    Exit Sub

cmdAceptar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCarp.cmdAceptar_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = vbKeyEscape) Then Unload Me

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCarp.Form_KeyPress", Erl)
    Resume Next
    
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoverForm Me.hwnd
End Sub

Private Sub List1_Click()
    
    On Error GoTo List1_Click_Err

    Dim SR As RECT, DR As RECT

    SR.Left = 0
    SR.Top = 0
    SR.Right = 32
    SR.Bottom = 32

    DR.Left = 0
    DR.Top = 0
    DR.Right = 32
    DR.Bottom = 32
    
    Call Grh_Render_To_Hdc(picture1, IIf(List1.ListIndex = 0, 550, 5348), 0, 0, False)

    
    Exit Sub

List1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCarp.List1_Click", Erl)
    Resume Next
    
End Sub

Private Sub lstArmas_Click()
    
    On Error GoTo lstArmas_Click_Err
    

    

    Dim SR As RECT, DR As RECT

    SR.Left = 0
    SR.Top = 0
    SR.Right = 32
    SR.Bottom = 32

    DR.Left = 0
    DR.Top = 0
    DR.Right = 32
    DR.Bottom = 32
    
    Call frmCarp.List1.Clear
    Call frmCarp.List2.Clear
    
    Call frmCarp.List1.AddItem("Leña")
    Call frmCarp.List1.AddItem("Leña elfica")
    Call frmCarp.List2.AddItem(ObjData(ObjCarpintero(lstArmas.ListIndex + 1)).Madera)
    Call frmCarp.List2.AddItem(ObjData(ObjCarpintero(lstArmas.ListIndex + 1)).MaderaElfica)

    desc.Caption = ObjData(ObjCarpintero(lstArmas.ListIndex + 1)).Texto

    Call Grh_Render_To_Hdc(Me.picture1, ObjData(ObjCarpintero(lstArmas.ListIndex + 1)).GrhIndex, 0, 0)
    
    picture1.Visible = True
    
    
    Exit Sub

lstArmas_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCarp.lstArmas_Click", Erl)
    Resume Next
    
End Sub
