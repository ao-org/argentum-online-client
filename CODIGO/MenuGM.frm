VERSION 5.00
Begin VB.Form MenuGM 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   387
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BAN"
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
      Index           =   15
      Left            =   840
      TabIndex        =   15
      Top             =   5475
      Width           =   330
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CARCEL"
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
      Index           =   14
      Left            =   660
      TabIndex        =   14
      Top             =   5100
      Width           =   630
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VIGILANTE"
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
      Index           =   13
      Left            =   360
      TabIndex        =   13
      Top             =   4755
      Width           =   1170
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   15
      Left            =   0
      Top             =   5400
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   14
      Left            =   0
      Top             =   5040
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   13
      Left            =   0
      Top             =   4680
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   12
      Left            =   0
      Top             =   4320
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SILENCIAR"
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
      Index           =   12
      Left            =   0
      TabIndex        =   12
      Top             =   4395
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   11
      Left            =   0
      Top             =   3960
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PENAS"
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
      Index           =   11
      Left            =   0
      TabIndex        =   11
      Top             =   4035
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   10
      Left            =   0
      Top             =   3600
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SEGUIR MOUSE"
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
      Index           =   10
      Left            =   0
      TabIndex        =   10
      Top             =   3675
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EJECUTAR"
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
      Index           =   9
      Left            =   0
      TabIndex        =   9
      Top             =   3315
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   9
      Left            =   0
      Top             =   3240
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ECHAR"
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
      Index           =   8
      Left            =   0
      TabIndex        =   8
      Top             =   2955
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   8
      Left            =   0
      Top             =   2880
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "REVIVIR"
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
      Index           =   7
      Left            =   0
      TabIndex        =   7
      Top             =   2595
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   7
      Left            =   0
      Top             =   2520
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INVENTARIO"
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
      Index           =   6
      Left            =   0
      TabIndex        =   6
      Top             =   2235
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   6
      Left            =   0
      Top             =   2160
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BOVEDA"
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
      Index           =   5
      Left            =   0
      TabIndex        =   5
      Top             =   1875
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   5
      Left            =   0
      Top             =   1800
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   4
      Left            =   0
      Top             =   1440
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BAL"
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
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Top             =   1515
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   3
      Left            =   0
      Top             =   1080
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "STAT"
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
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   1155
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   2
      Left            =   0
      Top             =   720
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INFO"
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
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   795
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   1
      Left            =   0
      Top             =   360
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CONSULTA"
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
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   435
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SUM"
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
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1950
   End
End
Attribute VB_Name = "MenuGM"
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

Private Over As Integer

Private Sub Form_Load()
    Call Aplicar_Transparencia(Me.hWnd, 180)
    OpcionLbl(0).Caption = JsonLanguage.Item("FORM_OPCION_9")
    OpcionLbl(1).Caption = JsonLanguage.Item("FORM_OPCION_10")
    OpcionLbl(2).Caption = JsonLanguage.Item("FORM_OPCION_11")
    OpcionLbl(3).Caption = JsonLanguage.Item("FORM_OPCION_12")
    OpcionLbl(4).Caption = JsonLanguage.Item("FORM_OPCION_13")
    OpcionLbl(5).Caption = JsonLanguage.Item("FORM_OPCION_14")
    OpcionLbl(6).Caption = JsonLanguage.Item("FORM_OPCION_15")
    OpcionLbl(7).Caption = JsonLanguage.Item("FORM_OPCION_16")
    OpcionLbl(8).Caption = JsonLanguage.Item("FORM_OPCION_17")
    OpcionLbl(9).Caption = JsonLanguage.Item("FORM_OPCION_18")
    OpcionLbl(10).Caption = JsonLanguage.Item("FORM_OPCION_19")
    OpcionLbl(11).Caption = JsonLanguage.Item("FORM_OPCION_20")
    OpcionLbl(12).Caption = JsonLanguage.Item("FORM_OPCION_21")
    OpcionLbl(13).Caption = JsonLanguage.Item("FORM_OPCION_22")
    OpcionLbl(14).Caption = JsonLanguage.Item("FORM_OPCION_23")
    OpcionLbl(15).Caption = JsonLanguage.Item("FORM_OPCION_24")

    
    Over = -1
End Sub

Private Sub OpcionImg_Click(Index As Integer)

    Dim tmp As String
    Dim tmptime As String

    Select Case Index

        Case 0
            Call ParseUserCommand("/SUM " & TargetName)

        Case 1
            Call ParseUserCommand("/CONSULTA " & TargetName)

        Case 2
            Call ParseUserCommand("/INFO " & TargetName)
            
        Case 3
            Call ParseUserCommand("/STAT " & TargetName)

        Case 4
            Call ParseUserCommand("/BAL " & TargetName)

        Case 5
            Call ParseUserCommand("/BOV " & TargetName)

        Case 6
            Call ParseUserCommand("/INV " & TargetName)
            
        Case 7
            Call ParseUserCommand("/revivir " & TargetName)

        Case 8
            Call ParseUserCommand("/echar " & TargetName)

        Case 9
            Call ParseUserCommand("/ejecutar " & TargetName)
'
        Case 10
            Call ParseUserCommand("/SM " & TargetName)

        Case 11
            Call ParseUserCommand("/PENAS " & TargetName)

        Case 12
            tmp = InputBox("Escriba el motivo del Silenciado.", "Silenciaso de " & TargetName)
            
            If tmp = "" Then
                InputBox ("No se puede Silenciar si dar motivos a " & TargetName)
            Else
                Call ParseUserCommand("/SILENCIAR " & TargetName & "@" & tmp)
            End If
            
        Case 13
            Dim mensajes(1 To 15) As String
                mensajes(1) = JsonLanguage.Item("MENSAJE_HORA_ACTUAL")
                mensajes(2) = JsonLanguage.Item("MENSAJE_LLUVIA")
                mensajes(3) = JsonLanguage.Item("MENSAJE_ZONA_SEGURA")
                mensajes(4) = JsonLanguage.Item("MENSAJE_ARBOLES_O_ARENA")
                mensajes(5) = JsonLanguage.Item("MENSAJE_DIA_O_NOCHE")
                mensajes(6) = JsonLanguage.Item("MENSAJE_CIELO_CLARO_O_OSCURO")
                mensajes(7) = JsonLanguage.Item("MENSAJE_ENTORNO_DESIERTO_O_BOSQUE")
                mensajes(8) = JsonLanguage.Item("MENSAJE_HORA_SERVIDOR")
                mensajes(9) = JsonLanguage.Item("MENSAJE_ZONA_SEGURA_PERSONAJE")
                mensajes(10) = JsonLanguage.Item("MENSAJE_LLUVIA_O_SECO")
                mensajes(11) = JsonLanguage.Item("MENSAJE_ENTORNO_BOSQUE_O_DESIERTO")
                mensajes(12) = JsonLanguage.Item("MENSAJE_MOMENTO_DEL_DIA")
                mensajes(13) = JsonLanguage.Item("MENSAJE_NOCHE_ACTUAL")
                mensajes(14) = JsonLanguage.Item("MENSAJE_ZONA_SEGURA_EN_PANTALLA")
                mensajes(15) = JsonLanguage.Item("MENSAJE_HORA_DEL_JUEGO")

            Dim MensajeSeleccionado As String
            Dim idx As Integer
            
            Randomize
            idx = Int((15 * Rnd) + 1)
            MensajeSeleccionado = Replace(mensajes(idx), "¬1", TargetName)
            
            Call ParseUserCommand("/MENSAJEINFORMACION " & TargetName & "@" & MensajeSeleccionado)
            ' agregar que el mensaje lo pueda leer yo tambien
            Call AddtoRichTextBox(frmMain.RecTxt, "MENSAJE A " & TargetName & ": " & MensajeSeleccionado, 0, 255, 255, True)


            
        Case 14
            tmp = InputBox("Escriba el motivo de Carcel .", "Carcel a " & TargetName)
            tmptime = InputBox("Escriba el tiempo de Carcel .", "Tiempo de Carcel a " & TargetName)
            If tmp = "" Or tmptime = "" Then
                MsgBox "Faltan datos. Repita la acción.", vbExclamation, "Error"
            Else
                Call WriteJail(TargetName, tmp, tmptime)
            End If
        Case 15
        
            Call ParseUserCommand("/BAN")
            tmp = InputBox("Escriba el motivo del BAN.", "Baneo de " & TargetName)

            If tmp = "" Then
                InputBox ("No se puede bannear si dar motivos a " & TargetName)
            Else
                Call WriteBanChar(TargetName, tmp)
            End If
            
    End Select

    Unload Me
    
End Sub

Private Sub OpcionImg_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Over <> Index Then
        If Over >= 0 Then
            OpcionLbl(Over).ForeColor = vbWhite
        End If
        OpcionLbl(Index).ForeColor = vbYellow
        Over = Index
    End If
End Sub

Private Sub OpcionLbl_Click(Index As Integer)
    Call OpcionImg_Click(Index)
End Sub

Private Sub OpcionLbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call OpcionImg_MouseMove(Index, Button, Shift, x, y)
End Sub

Public Sub LostFocus()
    If Over >= 0 Then
        OpcionLbl(Over).ForeColor = vbWhite
        Over = -1
    End If
End Sub
