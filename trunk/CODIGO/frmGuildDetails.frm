VERSION 5.00
Begin VB.Form frmGuildDetails 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fundar Clan"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   3345
   ClipControls    =   0   'False
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
   ScaleHeight     =   3960
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Nombre del clan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtClanName 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         MaxLength       =   30
         TabIndex        =   7
         Top             =   240
         Width           =   2895
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
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   555
         Width           =   2805
      End
   End
   Begin VB.Frame framAlign 
      Caption         =   "Alineamiento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   3135
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmGuildDetails.frx":0000
         Left            =   240
         List            =   "frmGuildDetails.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "El alineamiento del clan es el que decide qu� tipo de miembro podr� ingresar al clan y cu�l no."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fundar clan"
      Height          =   375
      Index           =   1
      Left            =   120
      MouseIcon       =   "frmGuildDetails.frx":0023
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3480
      Width           =   3120
   End
   Begin VB.Frame frmDesc 
      Caption         =   "Descripci�n"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3135
      Begin VB.TextBox txtDesc 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmGuildDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Private Sub Command1_Click(Index As Integer)

    Select Case Index

        Case 0
            Unload Me
        
        Case 1

            Dim fdesc      As String

            Dim Codex()    As String

            Dim k          As Byte

            Dim Cont       As Byte

            Dim Alineacion As Byte
            
            If txtClanName.Text = "" Then
                MensajeAdvertencia "�Ingrese un nombre!"
                Exit Sub

            End If

            If Len(txtClanName.Text) <= 30 Then
                If Not AsciiValidos(txtClanName) Then
                    MensajeAdvertencia "Nombre invalido."
                    Exit Sub

                End If

            Else
                MensajeAdvertencia "Nombre demasiado extenso."
                Exit Sub

            End If

            ClanName = txtClanName
    
            fdesc = Replace(txtDesc, vbCrLf, "�", , , vbBinaryCompare)
    
            '    If Not AsciiValidos(fdesc) Then
            '        MsgBox "La descripcion contiene caracteres invalidos"
            '        Exit Sub
            '    End If
            
            If Combo1.Text = "" Then
                MensajeAdvertencia "Debes definir el alineamiento del clan."
                Exit Sub

            End If
    
            If CreandoClan Then
                If Combo1.Text = "" Then
                    MensajeAdvertencia "Debes definir el alineamiento del clan."
                    Exit Sub

                End If

                If UCase$(Combo1.Text) = "CIUDADANA" Then
                    Alineacion = eClanType.ct_Legal
                ElseIf UCase$(Combo1.Text) = "CRIMINAL" Then
                    Alineacion = eClanType.ct_Evil

                End If
        
                Call WriteCreateNewGuild(fdesc, ClanName, Alineacion)
            Else
                Call WriteClanCodexUpdate(fdesc)

            End If

            CreandoClan = False
            Unload Me
            
    End Select

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
    Call FormParser.Parse_Form(Me)

End Sub

