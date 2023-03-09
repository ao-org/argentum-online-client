VERSION 5.00
Begin VB.Form MenuGM 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   316
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Caption         =   "SEGUIRMOUSE"
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
      Caption         =   "USO DE MACROS"
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
      Caption         =   "ADVERTENCIA"
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
      Caption         =   "BANEAR"
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
      Caption         =   "CARCEL 5 min"
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
      Caption         =   "INFO DE IP"
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
      Top             =   75
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
    Call Aplicar_Transparencia(Me.hwnd, 180)
    
    Over = -1
End Sub

Private Sub OpcionImg_Click(Index As Integer)

    Dim tmp As String

    Select Case Index

        Case 0
            Call ParseUserCommand("/SUM")

        Case 1
            Call ParseUserCommand("/SILENCIAR " & TargetName & "@" & "15")

        Case 2
            Call ParseUserCommand("/NICK2IP " & TargetName)
            Call WriteLastIP(TargetName)

        Case 3
            Call ParseUserCommand("/INFO " & TargetName)

        Case 4
            Call ParseUserCommand("/INV " & TargetName)

        Case 5
            Call ParseUserCommand("/REVIVIR " & TargetName)

        Case 6
            Call ParseUserCommand("/CONSULTA " & TargetName)

        Case 7
            Call ParseUserCommand("/CARCEL") ' ver ReyarB
            tmp = InputBox("Escriba el motivo de Carcel de 60 min.", "Carcel a " & TargetName)

            If tmp = "" Then
                InputBox ("No se puede mandar a carcel 60 min si dar motivos a " & TargetName)
            Else
                Call WriteJail(TargetName, tmp, "60")
            End If

        Case 8
            Call ParseUserCommand("/BAN") ' ver ReyarB
            tmp = InputBox("Escriba el motivo del BAN.", "Baneo de " & TargetName)

            If tmp = "" Then
                InputBox ("No se puede bannear si dar motivos a " & TargetName)
            Else
                Call WriteBanChar(TargetName, tmp)
            End If

        Case 9
            Call ParseUserCommand("/PENAS " & TargetName)

        Case 10
            tmp = InputBox("Escriba el motivo de la advertencia.", "Advertir a " & TargetName)
            Call ParseUserCommand("/ADVERTENCIA " & TargetName & "@" & tmp)

        Case 11
            Call WriteWhisper(TargetName, " El uso de macros o programas de tercero para automatizar estan prohibidos estos puede ocasionar sanciones que puede ir hasta Ban de IP.")

            'Call WriteWhisper(TargetName, " El Servidor comprobo que el cliente no es el Oficial por favor descargar de Steam Gracias.")
        Case 12
            Call ParseUserCommand("/SM " & TargetName)
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
