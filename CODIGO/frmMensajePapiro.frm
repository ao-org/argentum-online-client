VERSION 5.00
Begin VB.Form frmMensajePapiro 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Stream2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Plus De Ulla"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6720
      TabIndex        =   9
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Stream1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ao20 Oficial"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6720
      TabIndex        =   8
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label CafecitoLink 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ayudar con Cafecito"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   2
      Left            =   4560
      TabIndex        =   7
      Top             =   6225
      Width           =   1095
   End
   Begin VB.Label PatreonLink 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ayudar con Patreon"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   1
      Left            =   3285
      TabIndex        =   6
      Top             =   6225
      Width           =   1020
   End
   Begin VB.Label GameRulesLink 
      BackStyle       =   0  'Transparent
      Caption         =   "Ir al reglamento"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   3285
      TabIndex        =   5
      Top             =   5625
      Width           =   1455
   End
   Begin VB.Label HelpGuideLink 
      BackStyle       =   0  'Transparent
      Caption         =   "Ir a Guía"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3285
      TabIndex        =   4
      Top             =   5025
      Width           =   1335
   End
   Begin VB.Label BasicInfoLink 
      BackStyle       =   0  'Transparent
      Caption         =   "Ir al manual"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3285
      TabIndex        =   3
      Top             =   4420
      Width           =   5535
   End
   Begin VB.Label eventNews 
      BackStyle       =   0  'Transparent
      Caption         =   "Ver próximos eventos"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3285
      TabIndex        =   2
      Top             =   3825
      Width           =   1815
   End
   Begin VB.Label newsLink 
      BackStyle       =   0  'Transparent
      Caption         =   "Ir a novedades"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3285
      TabIndex        =   1
      Top             =   3225
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   8655
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   9840
      Top             =   720
      Width           =   495
   End
End
Attribute VB_Name = "frmMensajePapiro"
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

Private Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongW" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
    
Private Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongW" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
    
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal crKey As Long, _
    ByVal bAlpha As Byte, _
    ByVal dwFlags As Long) As Long
    
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1&

Private Sub BasicInfoLink_Click()
    Call OpenLink("https://ao20.com.ar/wiki")
End Sub

Private Sub CafecitoLink_Click(Index As Integer)
    Call OpenLink("https://cafecito.app/nolandstudios")
End Sub

Private Sub eventNews_Click()
    Call OpenLink("https://discord.com/channels/761213868352471040/843170070275686420")
End Sub

Private Sub Form_Load()

    Me.Picture = LoadInterface("board.bmp")
    MakeFormTransparent Me, vbBlack    'Set the Form "transparent by color."

End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub GameRulesLink_Click(Index As Integer)
    Call OpenLink("https://ao20.com.ar/reglamento")
End Sub

Private Sub HelpGuideLink_Click()
    Call OpenLink("https://ao20.com.ar/wiki/guia/comenzar")
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub


Private Sub newsLink_Click()
    Call OpenLink("https://steamcommunity.com/app/1956740/allnews/")
End Sub

Private Sub OpenLink(link As String)
    ShellExecute ByVal 0&, "open", _
        link, _
        vbNullString, vbNullString, _
        vbMaximizedFocus
End Sub

Private Sub PatreonLink_Click(Index As Integer)
    Call OpenLink("https://www.patreon.com/nolandstudios")
End Sub

Private Sub Stream1_Click()
    Call OpenLink("https://www.twitch.tv/ao20oficial")
End Sub

Private Sub Stream2_Click()
    Call OpenLink("https://www.twitch.tv/PlusAo20")
End Sub
