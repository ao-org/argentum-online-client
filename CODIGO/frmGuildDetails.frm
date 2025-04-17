VERSION 5.00
Begin VB.Form frmGuildDetails 
   BorderStyle     =   0  'None
   Caption         =   "Fundar Clan"
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   -180
   ClientWidth     =   3750
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
   ScaleHeight     =   4410
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00000000&
      ForeColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1600
      Width           =   3015
   End
   Begin VB.TextBox txtClanName 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   360
      MaxLength       =   30
      TabIndex        =   0
      Top             =   780
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      ItemData        =   "frmGuildDetails.frx":0000
      Left            =   360
      List            =   "frmGuildDetails.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2800
      Width           =   3015
   End
   Begin VB.Image cmdFundar 
      Height          =   495
      Left            =   850
      Top             =   3700
      Width           =   1695
   End
   Begin VB.Image cmdCerrar 
      Height          =   375
      Left            =   3280
      Top             =   5
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "El alineamiento del clan es el que decide qué tipo de miembro podrá ingresar al clan y cuál no."
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
      Height          =   855
      Left            =   360
      TabIndex        =   4
      Top             =   3175
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nota: No se toleraran nombres inapropiados."
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
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1090
      Width           =   3045
   End
End
Attribute VB_Name = "frmGuildDetails"
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
Private cBotonFundar As clsGraphicalButton

Private Sub loadButtons()
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonFundar = New clsGraphicalButton
    
    Call cBotonCerrar.Initialize(cmdCerrar, "boton-cerrar-default.bmp", _
                                                "boton-cerrar-over.bmp", _
                                                "boton-cerrar-off.bmp", Me)
                                                
    Call cBotonFundar.Initialize(cmdFundar, "boton-fundar-clan-default.bmp", _
                                                    "boton-fundar-clan-over.bmp", _
                                                    "boton-fundar-clan-off.bmp", Me)
End Sub
Private Sub cmdcerrar_Click()
    Unload Me
End Sub

Private Sub cmdFundar_Click()

    
    On Error GoTo cmdFundar_Click_Err
    
            Dim fdesc      As String

            Dim Codex()    As String

            Dim k          As Byte

            Dim Cont       As Byte

            Dim Alineacion As Byte
            
            If txtClanName.Text = "" Then
                MensajeAdvertencia JsonLanguage.Item("ADVERTENCIA_INGRESAR_NOMBRE")
                
                Exit Sub

            End If

            If Len(txtClanName.Text) <= 30 Then
                If Not AsciiValidos(txtClanName) Then
                    MensajeAdvertencia JsonLanguage.Item("ADVERTENCIA_NOMBRE_INVALIDO")
                    Exit Sub

                End If

            Else
                MensajeAdvertencia JsonLanguage.Item("ADVERTENCIA_NOMBRE_DEMASIADO_EXTENSO")
                Exit Sub

            End If

            ClanName = txtClanName
    
            fdesc = Replace(txtDesc, vbCrLf, "º", , , vbBinaryCompare)
    
            If Combo1.Text = "" Then
                MensajeAdvertencia JsonLanguage.Item("ADVERTENCIA_DEFINIR_ALINEAMIENTO_CLAN")
                Exit Sub

            End If
    
            If CreandoClan Then
                If Combo1.ListIndex < 0 Then
                    MensajeAdvertencia JsonLanguage.Item("ADVERTENCIA_DEFINIR_ALINEAMIENTO_CLAN")
                    Exit Sub
                End If

                If Combo1.ListIndex = eClanType.ct_Neutral Then
                    Alineacion = eClanType.ct_Neutral
                ElseIf Combo1.ListIndex = eClanType.ct_Real Then
                    Alineacion = eClanType.ct_Real
                ElseIf Combo1.ListIndex = eClanType.ct_Caos Then
                    Alineacion = eClanType.ct_Caos
                ElseIf Combo1.ListIndex = eClanType.ct_Ciudadana Then
                    Alineacion = eClanType.ct_Ciudadana
                ElseIf Combo1.ListIndex = eClanType.ct_Criminal Then
                    Alineacion = eClanType.ct_Criminal
                End If
        
                Call WriteCreateNewGuild(fdesc, ClanName, Alineacion)
            Else
                Call WriteClanCodexUpdate(fdesc)

            End If

            CreandoClan = False
            Unload Me
               
    Exit Sub

cmdFundar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmGuildDetails.cmdFundar_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_Deactivate()

    'If Not frmGuildLeader.Visible Then
    '    Me.SetFocus
    'Else
    '    'Unload Me
    'End If
    '
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    Me.Picture = LoadInterface("ventanaclanes_fundar.bmp")
    Call Aplicar_Transparencia(Me.hwnd, 240)
    Call loadButtons
    
    Exit Sub


Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildDetails.Form_Load", Erl)
    Resume Next
    
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    Call MoverForm(Me.hwnd)
    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmGuildDetails.Form_MouseMove", Erl)
    Resume Next
    
End Sub
