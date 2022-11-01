VERSION 5.00
Begin VB.Form frmHerrero 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Trabajar con Herreria"
   ClientHeight    =   6528
   ClientLeft      =   0
   ClientTop       =   -96
   ClientWidth     =   7056
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   544
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   588
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List2 
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
      Left            =   5760
      TabIndex        =   5
      Top             =   2520
      Width           =   645
   End
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
      Height          =   2880
      Left            =   480
      TabIndex        =   4
      Top             =   2325
      Width           =   2775
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
      Left            =   5265
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "1"
      Top             =   4275
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   4890
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   1845
      Width           =   480
   End
   Begin VB.Image cmdMenos 
      Height          =   315
      Left            =   4740
      Tag             =   "0"
      Top             =   4215
      Width           =   315
   End
   Begin VB.Image cmdMas 
      Height          =   315
      Left            =   6000
      Tag             =   "0"
      Top             =   4215
      Width           =   315
   End
   Begin VB.Image cmdCerrar 
      Height          =   420
      Left            =   6600
      Top             =   0
      Width           =   420
   End
   Begin VB.Label desc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   3900
      Width           =   2535
   End
   Begin VB.Image cmdConstruir 
      Height          =   420
      Left            =   4155
      Top             =   4680
      Width           =   1980
   End
   Begin VB.Image cmdAceptar 
      Height          =   420
      Left            =   2535
      Top             =   5730
      Width           =   1980
   End
   Begin VB.Image Command2 
      Height          =   510
      Left            =   2595
      Stretch         =   -1  'True
      Top             =   1365
      Width           =   570
   End
   Begin VB.Image Command1 
      Height          =   600
      Left            =   1965
      OLEDropMode     =   1  'Manual
      Top             =   1350
      Width           =   615
   End
   Begin VB.Image Command4 
      Height          =   600
      Left            =   1305
      Tag             =   "0"
      Top             =   1365
      Width           =   615
   End
   Begin VB.Image Command3 
      Height          =   600
      Left            =   660
      Top             =   1350
      Width           =   615
   End
End
Attribute VB_Name = "frmHerrero"
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

'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
Dim Index As Byte

Option Explicit

Private cBotonAceptar As clsGraphicalButton
Private cBotonConstruir As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton
Private cBotonMas As clsGraphicalButton
Private cBotonMenos As clsGraphicalButton


Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call Aplicar_Transparencia(Me.hwnd, 240)
    
    Call FormParser.Parse_Form(Me)
    
    Me.Picture = LoadInterface("VentanaHerreria.bmp")
    Call LoadButtons
    
    Index = 3
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmHerrero.Form_Load", Erl)
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
Private Sub Command1_Click()
    
    On Error GoTo Command1_Click_Err
    
    Index = 1

    Dim i As Byte

    lstArmas.Clear

    For i = 1 To UBound(ArmasHerrero())

        If ArmasHerrero(i).Index = 0 Then Exit For
        Call frmHerrero.lstArmas.AddItem(ObjData(ArmasHerrero(i).Index).Name)
    Next i
    
    Command1.Picture = LoadInterface("boton-espada-over.bmp")
    Command3.Picture = Nothing
    Command2.Picture = Nothing
    Command4.Picture = Nothing

    
    Exit Sub

Command1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmHerrero.Command1_Click", Erl)
    Resume Next
    
End Sub


Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Index <> 1 Then Command1.Picture = LoadInterface("boton-espada-off.bmp")
End Sub

Private Sub Command2_Click()
    
    On Error GoTo Command2_Click_Err
    
    Index = 2

    Dim i As Byte

    lstArmas.Clear

    For i = 0 To UBound(ArmadurasHerrero())

        If ArmadurasHerrero(i).Index = 0 Then Exit For
        If ObjData(ArmadurasHerrero(i).Index).ObjType = 3 Then
            Call frmHerrero.lstArmas.AddItem(ObjData(ArmadurasHerrero(i).Index).Name)

        End If

    Next i
    
    Command2.Picture = LoadInterface("boton-armadura-over.bmp")
    Command1.Picture = Nothing
    Command3.Picture = Nothing
    Command4.Picture = Nothing

    
    Exit Sub

Command2_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmHerrero.Command2_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index <> 2 Then Command2.Picture = LoadInterface("boton-armadura-off.bmp")
End Sub

Private Sub Command3_Click()
    
    On Error GoTo Command3_Click_Err
    
    Index = 3

    lstArmas.Clear

    Dim i As Byte

    For i = 0 To UBound(CascosHerrero())

        If CascosHerrero(i).Index = 0 Then Exit For
        Call frmHerrero.lstArmas.AddItem(ObjData(CascosHerrero(i).Index).Name)
    Next i

    Command3.Picture = LoadInterface("boton-casco-over.bmp")
    Command1.Picture = Nothing
    Command2.Picture = Nothing
    Command4.Picture = Nothing

    
    Exit Sub

Command3_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmHerrero.Command3_Click", Erl)
    Resume Next
    
End Sub


Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index <> 3 Then Command3.Picture = LoadInterface("boton-casco-off.bmp")
End Sub

Private Sub Command4_Click()
    
    On Error GoTo Command4_Click_Err
    
    Index = 4

    lstArmas.Clear

    Dim i As Byte

    For i = 0 To UBound(EscudosHerrero())

        If EscudosHerrero(i).Index = 0 Then Exit For
        Call frmHerrero.lstArmas.AddItem(ObjData(EscudosHerrero(i).Index).Name)
    Next i

    Command4.Picture = LoadInterface("boton-escudo-over.bmp")
    Command1.Picture = Nothing
    Command2.Picture = Nothing
    Command3.Picture = Nothing

    
    Exit Sub

Command4_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmHerrero.Command4_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index <> 4 Then Command4.Picture = LoadInterface("boton-escudo-off.bmp")
End Sub

Private Sub cmdAceptar_Click()
    
    On Error GoTo cmdAceptar_Click_Err
    
    Unload Me

    
    Exit Sub

cmdAceptar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmHerrero.cmdAceptar_Click", Erl)
    Resume Next
    
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

    If Index = 1 Then

        If cantidad > 1 Then
            UserMacro.cantidad = cantidad
            UserMacro.TIPO = 4
            UserMacro.Index = ArmasHerrero(lstArmas.ListIndex + 1).Index
            AddtoRichTextBox frmMain.RecTxt, "Comienzas a trabajar.", 2, 51, 223, 1, 1
            UserMacro.Intervalo = IntervaloTrabajoConstruir
            UserMacro.Activado = True
            frmMain.MacroLadder.Interval = IntervaloTrabajoConstruir
            frmMain.MacroLadder.Enabled = True
        Else
            Call WriteCraftBlacksmith(ArmasHerrero(lstArmas.ListIndex + 1).Index)

            If frmMain.macrotrabajo.Enabled Then MacroBltIndex = ArmasHerrero(lstArmas.ListIndex + 1).Index

        End If

        Unload Me
    ElseIf Index = 2 Then
    
        If cantidad > 1 Then
            UserMacro.cantidad = cantidad
            UserMacro.TIPO = 4
            UserMacro.Index = ArmadurasHerrero(lstArmas.ListIndex).Index
            AddtoRichTextBox frmMain.RecTxt, "Comienzas a trabajar.", 2, 51, 223, 1, 1
            UserMacro.Intervalo = IntervaloTrabajoConstruir
            UserMacro.Activado = True
            frmMain.MacroLadder.Interval = IntervaloTrabajoConstruir
            frmMain.MacroLadder.Enabled = True
        Else
            Call WriteCraftBlacksmith(ArmadurasHerrero(lstArmas.ListIndex).Index)

            If frmMain.macrotrabajo.Enabled Then MacroBltIndex = ArmadurasHerrero(lstArmas.ListIndex).Index
            
        End If
        
        Unload Me
    ElseIf Index = 3 Then
    
        If cantidad > 1 Then
            UserMacro.cantidad = cantidad
            UserMacro.TIPO = 4
            UserMacro.Index = CascosHerrero(lstArmas.ListIndex).Index
            AddtoRichTextBox frmMain.RecTxt, "Comienzas a trabajar.", 2, 51, 223, 1, 1
            UserMacro.Intervalo = IntervaloTrabajoConstruir
            UserMacro.Activado = True
            frmMain.MacroLadder.Interval = IntervaloTrabajoConstruir
            frmMain.MacroLadder.Enabled = True
        Else
            Call WriteCraftBlacksmith(CascosHerrero(lstArmas.ListIndex).Index)

            If frmMain.macrotrabajo.Enabled Then MacroBltIndex = CascosHerrero(lstArmas.ListIndex).Index
            
        End If
        
        Unload Me
    ElseIf Index = 4 Then
    
        If cantidad > 1 Then
            UserMacro.cantidad = cantidad
            UserMacro.TIPO = 4
            UserMacro.Index = EscudosHerrero(lstArmas.ListIndex).Index
            AddtoRichTextBox frmMain.RecTxt, "Comienzas a trabajar.", 2, 51, 223, 1, 1
            UserMacro.Intervalo = IntervaloTrabajoConstruir
            UserMacro.Activado = True
            frmMain.MacroLadder.Interval = IntervaloTrabajoConstruir
            frmMain.MacroLadder.Enabled = True
        Else
            Call WriteCraftBlacksmith(EscudosHerrero(lstArmas.ListIndex).Index)

            If frmMain.macrotrabajo.Enabled Then MacroBltIndex = EscudosHerrero(lstArmas.ListIndex).Index
            
        End If
        
        Unload Me

    End If

    
    Exit Sub

cmdConstruir_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmHerrero.cmdConstruir_Click", Erl)
    Resume Next
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err

    If (KeyAscii = 27) Then
        Unload Me
    End If
    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.number, Err.Description, "frmHerrero.Form_KeyPress", Erl)
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

    Dim grh As Long

    If List1.ListIndex = 0 Then
        grh = 724
    ElseIf List1.ListIndex = 1 Then
        grh = 725
    ElseIf List1.ListIndex = 2 Then
        grh = 723

    End If

    Call Grh_Render_To_Hdc(picture1, grh, 0, 0, False)

    
    Exit Sub

List1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmHerrero.List1_Click", Erl)
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

    List1.Clear
    List2.Clear
    List1.AddItem ("Lingote de hierro")
    List1.AddItem ("Lingo de plata")
    List1.AddItem ("Lingotes de Oro")

    If Index = 1 Then
        Call Grh_Render_To_Hdc(picture1, ObjData(ArmasHerrero(lstArmas.ListIndex + 1).Index).GrhIndex, 0, 0)
        List2.AddItem (ObjData(ArmasHerrero(lstArmas.ListIndex + 1).Index).LingH)
        List2.AddItem (ObjData(ArmasHerrero(lstArmas.ListIndex + 1).Index).LingP)
        List2.AddItem (ObjData(ArmasHerrero(lstArmas.ListIndex + 1).Index).LingO)
        desc.Caption = "Golpe: " & ObjData(ArmasHerrero(lstArmas.ListIndex + 1).Index).MinHit & "/" & ObjData(ArmasHerrero(lstArmas.ListIndex + 1).Index).MaxHit
    
    ElseIf Index = 2 Then
        Call Grh_Render_To_Hdc(picture1, ObjData(ArmadurasHerrero(lstArmas.ListIndex).Index).GrhIndex, 0, 0)
        List2.AddItem (ObjData(ArmadurasHerrero(lstArmas.ListIndex).Index).LingH)
        List2.AddItem (ObjData(ArmadurasHerrero(lstArmas.ListIndex).Index).LingP)
        List2.AddItem (ObjData(ArmadurasHerrero(lstArmas.ListIndex).Index).LingO)
        desc.Caption = "Defensa: " & ObjData(ArmadurasHerrero(lstArmas.ListIndex).Index).MinDef & "/" & ObjData(ArmadurasHerrero(lstArmas.ListIndex).Index).MaxDef
    ElseIf Index = 3 Then
        Call Grh_Render_To_Hdc(picture1, ObjData(CascosHerrero(lstArmas.ListIndex).Index).GrhIndex, 0, 0)
        List2.AddItem (ObjData(CascosHerrero(lstArmas.ListIndex).Index).LingH)
        List2.AddItem (ObjData(CascosHerrero(lstArmas.ListIndex).Index).LingP)
        List2.AddItem (ObjData(CascosHerrero(lstArmas.ListIndex).Index).LingO)
        desc.Caption = "Defensa: " & ObjData(CascosHerrero(lstArmas.ListIndex).Index).MinDef & "/" & ObjData(CascosHerrero(lstArmas.ListIndex).Index).MaxDef
    ElseIf Index = 4 Then
        Call Grh_Render_To_Hdc(picture1, ObjData(EscudosHerrero(lstArmas.ListIndex).Index).GrhIndex, 0, 0)
        List2.AddItem (ObjData(EscudosHerrero(lstArmas.ListIndex).Index).LingH)
        List2.AddItem (ObjData(EscudosHerrero(lstArmas.ListIndex).Index).LingP)
        List2.AddItem (ObjData(EscudosHerrero(lstArmas.ListIndex).Index).LingO)
        desc.Caption = "Defensa: " & ObjData(EscudosHerrero(lstArmas.ListIndex).Index).MinDef & "/" & ObjData(EscudosHerrero(lstArmas.ListIndex).Index).MaxDef

    End If

    picture1.Visible = True

    
    Exit Sub

lstArmas_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmHerrero.lstArmas_Click", Erl)
    Resume Next
    
End Sub

